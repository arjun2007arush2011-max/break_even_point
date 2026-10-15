#!/usr/bin/env ruby
# ==============================================================================
# Break-Even Point Predictor - Ruby Backend Server (WEBrick + SQLite3)
# Zero external dependencies; uses Ruby standard library + native sqlite3.
# Serves the frontend at http://localhost:8000/ and handles all REST API routes.
# ==============================================================================

require 'webrick'
require 'json'
require 'sqlite3'
require 'digest'
require 'securerandom'
require 'time'

PORT = (ENV['PORT'] || 8000).to_i
DB_PATH = File.expand_path('breakeven.db', __dir__)

# ------------------------------------------------------------------------------
# Database Initialization
# ------------------------------------------------------------------------------
def init_db(db_path)
  db = SQLite3::Database.new(db_path)
  db.results_as_hash = true

  db.execute <<-SQL
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      email TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      token TEXT,
      created_at TEXT NOT NULL
    );
  SQL

  db.execute <<-SQL
    CREATE TABLE IF NOT EXISTS profiles (
      user_id INTEGER PRIMARY KEY,
      firm_name TEXT,
      owner_name TEXT,
      business_type TEXT,
      phone TEXT,
      email TEXT,
      address TEXT,
      updated_at TEXT,
      FOREIGN KEY(user_id) REFERENCES users(id)
    );
  SQL

  db.execute <<-SQL
    CREATE TABLE IF NOT EXISTS inputs (
      user_id INTEGER PRIMARY KEY,
      fixed_costs REAL,
      selling_price REAL,
      variable_cost REAL,
      expected_volume REAL,
      target_profit REAL,
      currency TEXT,
      unit_label TEXT,
      updated_at TEXT,
      FOREIGN KEY(user_id) REFERENCES users(id)
    );
  SQL

  # Seed sample demo account if not exists
  existing = db.get_first_row("SELECT id FROM users WHERE email = 'demo@business.com'")
  unless existing
    demo_pass_hash = Digest::SHA256.hexdigest("Demo1234")
    now = Time.now.utc.iso8601
    db.execute(
      "INSERT INTO users (name, email, password_hash, created_at) VALUES (?, ?, ?, ?)",
      ['Elena Rostova', 'demo@business.com', demo_pass_hash, now]
    )
    user_id = db.last_insert_row_id

    db.execute(
      "INSERT INTO profiles (user_id, firm_name, owner_name, business_type, phone, email, address, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
      [user_id, 'Velvet Bean Roastery', 'Elena Rostova', 'Restaurant / Food', '2065550184', 'elena@velvetbean.com', 'Pike Place Market, Seattle, WA', now]
    )

    db.execute(
      "INSERT INTO inputs (user_id, fixed_costs, selling_price, variable_cost, expected_volume, target_profit, currency, unit_label, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
      [user_id, 12500.0, 6.50, 2.10, 3800.0, 5000.0, '$', 'cups', now]
    )
  end

  db
end

DB = init_db(DB_PATH)

# ------------------------------------------------------------------------------
# Mathematical Model Engine (Pure Ruby calculation)
# ------------------------------------------------------------------------------
def calculate_model(f, p, v, q, t)
  f = [0.0, f.to_f].max
  p = [0.0, p.to_f].max
  v = [0.0, v.to_f].max
  q = [0.0, q.to_f].max
  t = [0.0, t.to_f].max

  cm = p - v
  cm_ratio = p > 0 ? (cm / p) : 0.0
  is_valid = cm > 0

  be_units = is_valid ? (f / cm) : nil
  be_revenue = is_valid ? (be_units * p) : nil

  exp_revenue = p * q
  exp_variable_cost = v * q
  exp_total_cost = f + exp_variable_cost
  exp_profit = exp_revenue - exp_total_cost

  mos_units = is_valid ? (q - be_units) : nil
  mos_ratio = (is_valid && q > 0) ? ((q - be_units) / q) : 0.0
  mos_revenue = is_valid ? (mos_units * p) : nil

  target_units = is_valid ? ((f + t) / cm) : nil
  target_revenue = is_valid ? (target_units * p) : nil

  total_cm = q * cm
  dol = (exp_profit > 0) ? (total_cm / exp_profit) : nil

  {
    fixed_costs: f,
    selling_price: p,
    variable_cost: v,
    expected_volume: q,
    target_profit: t,
    contribution_margin: cm.round(2),
    contribution_margin_ratio: (cm_ratio * 100).round(2),
    is_valid: is_valid,
    break_even_units: be_units ? be_units.round(2) : nil,
    break_even_revenue: be_revenue ? be_revenue.round(2) : nil,
    expected_revenue: exp_revenue.round(2),
    expected_variable_cost: exp_variable_cost.round(2),
    expected_total_cost: exp_total_cost.round(2),
    expected_profit: exp_profit.round(2),
    margin_of_safety_units: mos_units ? mos_units.round(2) : nil,
    margin_of_safety_ratio: (mos_ratio * 100).round(2),
    margin_of_safety_revenue: mos_revenue ? mos_revenue.round(2) : nil,
    target_units: target_units ? target_units.round(2) : nil,
    target_revenue: target_revenue ? target_revenue.round(2) : nil,
    degree_of_operating_leverage: dol ? dol.round(2) : nil
  }
end

# ------------------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------------------
def set_cors_headers(res)
  res['Access-Control-Allow-Origin'] = '*'
  res['Access-Control-Allow-Methods'] = 'GET, POST, PUT, DELETE, OPTIONS'
  res['Access-Control-Allow-Headers'] = 'Content-Type, Authorization, X-Requested-With'
end

def json_response(res, status_code, data)
  set_cors_headers(res)
  res.status = status_code
  res['Content-Type'] = 'application/json'
  res.body = JSON.generate(data)
end

def parse_json(req)
  begin
    return JSON.parse(req.body) if req.body && !req.body.empty?
  rescue => e
    return {}
  end
  {}
end

def authenticate_user(req)
  auth_header = req['Authorization'] || ''
  token = auth_header.sub(/^Bearer\s+/, '').strip
  return nil if token.empty?

  DB.get_first_row("SELECT * FROM users WHERE token = ?", [token])
end

# ------------------------------------------------------------------------------
# HTTP Server & Endpoints
# ------------------------------------------------------------------------------
server = WEBrick::HTTPServer.new(
  Port: PORT,
  DocumentRoot: File.expand_path(__dir__),
  Logger: WEBrick::Log.new($stdout, WEBrick::Log::INFO),
  AccessLog: [[$stdout, WEBrick::AccessLog::COMBINED_LOG_FORMAT]]
)

# Handle preflight OPTIONS
server.mount_proc('/api') do |req, res|
  set_cors_headers(res)
  if req.request_method == 'OPTIONS'
    res.status = 200
    res.body = ''
  end
end

# 1. Health check
server.mount_proc('/api/health') do |req, res|
  set_cors_headers(res)
  return if req.request_method == 'OPTIONS'
  json_response(res, 200, {
    status: 'ok',
    backend: 'Ruby WEBrick + SQLite3',
    timestamp: Time.now.utc.iso8601,
    database: File.basename(DB_PATH)
  })
end

# 2. Authentication: Sign Up
server.mount_proc('/api/auth/signup') do |req, res|
  set_cors_headers(res)
  return if req.request_method == 'OPTIONS'
  unless req.request_method == 'POST'
    return json_response(res, 405, { error: 'Method not allowed' })
  end

  data = parse_json(req)
  name = (data['name'] || '').strip
  email = (data['email'] || '').strip.downcase
  password = data['password'] || ''
  pass_hash = data['password_hash'] || (password.empty? ? nil : Digest::SHA256.hexdigest(password))

  if name.empty? || email.empty? || !pass_hash
    return json_response(res, 400, { error: 'Name, email, and password are required.' })
  end

  existing = DB.get_first_row("SELECT id FROM users WHERE email = ?", [email])
  if existing
    return json_response(res, 409, { error: 'An account with this email already exists.' })
  end

  token = SecureRandom.hex(24)
  now = Time.now.utc.iso8601

  DB.execute(
    "INSERT INTO users (name, email, password_hash, token, created_at) VALUES (?, ?, ?, ?, ?)",
    [name, email, pass_hash, token, now]
  )
  user_id = DB.last_insert_row_id

  # Create default profile
  prof = data['profile'] || {}
  firm = prof['firm_name'] || "#{name}'s Business"
  b_type = prof['business_type'] || 'Retail'
  phone = prof['phone'] || '5551234567'
  prof_email = prof['email'] || email
  addr = prof['address'] || 'Downtown Headquarters'

  DB.execute(
    "INSERT INTO profiles (user_id, firm_name, owner_name, business_type, phone, email, address, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
    [user_id, firm, name, b_type, phone, prof_email, addr, now]
  )

  # Create default inputs
  inp = data['inputs'] || {}
  f = (inp['fixed_costs'] || 12000).to_f
  p = (inp['selling_price'] || 40).to_f
  v = (inp['variable_cost'] || 15).to_f
  q = (inp['expected_volume'] || 600).to_f
  t = (inp['target_profit'] || 3000).to_f
  curr = inp['currency'] || '$'
  unit = inp['unit_label'] || 'units'

  DB.execute(
    "INSERT INTO inputs (user_id, fixed_costs, selling_price, variable_cost, expected_volume, target_profit, currency, unit_label, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
    [user_id, f, p, v, q, t, curr, unit, now]
  )

  json_response(res, 201, {
    success: true,
    token: token,
    user: { id: user_id, name: name, email: email },
    profile: {
      firm_name: firm,
      owner_name: name,
      business_type: b_type,
      phone: phone,
      email: prof_email,
      address: addr
    },
    inputs: {
      fixed_costs: f,
      selling_price: p,
      variable_cost: v,
      expected_volume: q,
      target_profit: t,
      currency: curr,
      unit_label: unit
    }
  })
end

# 3. Authentication: Log In
server.mount_proc('/api/auth/login') do |req, res|
  set_cors_headers(res)
  return if req.request_method == 'OPTIONS'
  unless req.request_method == 'POST'
    return json_response(res, 405, { error: 'Method not allowed' })
  end

  data = parse_json(req)
  email = (data['email'] || '').strip.downcase
  password = data['password'] || ''
  pass_hash = data['password_hash'] || (password.empty? ? nil : Digest::SHA256.hexdigest(password))

  user = DB.get_first_row("SELECT * FROM users WHERE email = ?", [email])
  if !user || user['password_hash'] != pass_hash
    return json_response(res, 401, { error: 'Invalid email or password.' })
  end

  token = SecureRandom.hex(24)
  DB.execute("UPDATE users SET token = ? WHERE id = ?", [token, user['id']])

  profile = DB.get_first_row("SELECT * FROM profiles WHERE user_id = ?", [user['id']]) || {}
  inputs = DB.get_first_row("SELECT * FROM inputs WHERE user_id = ?", [user['id']]) || {}

  json_response(res, 200, {
    success: true,
    token: token,
    user: { id: user['id'], name: user['name'], email: user['email'] },
    profile: {
      firm_name: profile['firm_name'],
      owner_name: profile['owner_name'],
      business_type: profile['business_type'],
      phone: profile['phone'],
      email: profile['email'],
      address: profile['address']
    },
    inputs: {
      fixed_costs: inputs['fixed_costs'],
      selling_price: inputs['selling_price'],
      variable_cost: inputs['variable_cost'],
      expected_volume: inputs['expected_volume'],
      target_profit: inputs['target_profit'],
      currency: inputs['currency'] || '$',
      unit_label: inputs['unit_label'] || 'units'
    }
  })
end

# 4. Authentication: Current User Status & Session
server.mount_proc('/api/auth/me') do |req, res|
  set_cors_headers(res)
  return if req.request_method == 'OPTIONS'

  user = authenticate_user(req)
  unless user
    return json_response(res, 401, { error: 'Unauthorized', authenticated: false })
  end

  profile = DB.get_first_row("SELECT * FROM profiles WHERE user_id = ?", [user['id']]) || {}
  inputs = DB.get_first_row("SELECT * FROM inputs WHERE user_id = ?", [user['id']]) || {}

  json_response(res, 200, {
    authenticated: true,
    user: { id: user['id'], name: user['name'], email: user['email'] },
    profile: {
      firm_name: profile['firm_name'],
      owner_name: profile['owner_name'],
      business_type: profile['business_type'],
      phone: profile['phone'],
      email: profile['email'],
      address: profile['address']
    },
    inputs: {
      fixed_costs: inputs['fixed_costs'],
      selling_price: inputs['selling_price'],
      variable_cost: inputs['variable_cost'],
      expected_volume: inputs['expected_volume'],
      target_profit: inputs['target_profit'],
      currency: inputs['currency'] || '$',
      unit_label: inputs['unit_label'] || 'units'
    }
  })
end

# 5. Authentication: Logout
server.mount_proc('/api/auth/logout') do |req, res|
  set_cors_headers(res)
  return if req.request_method == 'OPTIONS'

  user = authenticate_user(req)
  if user
    DB.execute("UPDATE users SET token = NULL WHERE id = ?", [user['id']])
  end
  json_response(res, 200, { success: true, message: 'Logged out successfully' })
end

# 6. Business Profile: Get & Update
server.mount_proc('/api/profile') do |req, res|
  set_cors_headers(res)
  return if req.request_method == 'OPTIONS'

  user = authenticate_user(req)
  unless user
    return json_response(res, 401, { error: 'Unauthorized' })
  end

  if req.request_method == 'GET'
    profile = DB.get_first_row("SELECT * FROM profiles WHERE user_id = ?", [user['id']]) || {}
    json_response(res, 200, { success: true, profile: profile })
  elsif req.request_method == 'POST' || req.request_method == 'PUT'
    data = parse_json(req)
    now = Time.now.utc.iso8601

    DB.execute(<<-SQL, [data['firm_name'], data['owner_name'], data['business_type'], data['phone'], data['email'], data['address'], now, user['id']])
      UPDATE profiles
      SET firm_name = ?, owner_name = ?, business_type = ?, phone = ?, email = ?, address = ?, updated_at = ?
      WHERE user_id = ?;
    SQL

    updated = DB.get_first_row("SELECT * FROM profiles WHERE user_id = ?", [user['id']])
    json_response(res, 200, { success: true, profile: updated })
  else
    json_response(res, 405, { error: 'Method not allowed' })
  end
end

# 7. Inputs: Get & Update
server.mount_proc('/api/inputs') do |req, res|
  set_cors_headers(res)
  return if req.request_method == 'OPTIONS'

  user = authenticate_user(req)
  unless user
    return json_response(res, 401, { error: 'Unauthorized' })
  end

  if req.request_method == 'GET'
    inputs = DB.get_first_row("SELECT * FROM inputs WHERE user_id = ?", [user['id']]) || {}
    json_response(res, 200, { success: true, inputs: inputs })
  elsif req.request_method == 'POST' || req.request_method == 'PUT'
    data = parse_json(req)
    now = Time.now.utc.iso8601

    f = data['fixed_costs'].to_f
    p = data['selling_price'].to_f
    v = data['variable_cost'].to_f
    q = data['expected_volume'].to_f
    t = data['target_profit'].to_f
    curr = data['currency'] || '$'
    unit = data['unit_label'] || 'units'

    DB.execute(<<-SQL, [f, p, v, q, t, curr, unit, now, user['id']])
      UPDATE inputs
      SET fixed_costs = ?, selling_price = ?, variable_cost = ?, expected_volume = ?, target_profit = ?, currency = ?, unit_label = ?, updated_at = ?
      WHERE user_id = ?;
    SQL

    updated = DB.get_first_row("SELECT * FROM inputs WHERE user_id = ?", [user['id']])
    json_response(res, 200, { success: true, inputs: updated })
  else
    json_response(res, 405, { error: 'Method not allowed' })
  end
end

# 8. Server-side Model Calculation
server.mount_proc('/api/calculate') do |req, res|
  set_cors_headers(res)
  return if req.request_method == 'OPTIONS'

  unless req.request_method == 'POST'
    return json_response(res, 405, { error: 'Method not allowed' })
  end

  data = parse_json(req)
  f = data['fixed_costs'] || 0
  p = data['selling_price'] || 0
  v = data['variable_cost'] || 0
  q = data['expected_volume'] || 0
  t = data['target_profit'] || 0

  calc = calculate_model(f, p, v, q, t)
  json_response(res, 200, { success: true, calculation: calc })
end

# ------------------------------------------------------------------------------
# Start Server
# ------------------------------------------------------------------------------
trap('INT') { server.shutdown }
trap('TERM') { server.shutdown }

puts "\n======================================================="
puts " Break-Even Point Predictor - Backend Server Active"
puts " Local URL: http://localhost:#{PORT}/"
puts " SQLite DB: #{DB_PATH}"
puts " Press Ctrl+C to stop the server"
puts "=======================================================\n\n"

server.start
