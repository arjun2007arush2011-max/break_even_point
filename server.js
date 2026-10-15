// ==============================================================================
// Break-Even Point Predictor - Node.js Backend Server
// Zero external dependencies; uses Node.js standard modules (http, fs, path, crypto).
// Serves static frontend assets and full REST API routes with JSON file storage.
// ==============================================================================

var http = require('http');
var fs = require('fs');
var path = require('path');
var crypto = require('crypto');

var PORT = parseInt(process.env.PORT, 10) || 8000;
var DB_FILE = path.join(__dirname, 'database.json');

// Initialize database file
function initDb() {
  if (!fs.existsSync(DB_FILE)) {
    var demoHash = crypto.createHash('sha256').update('Demo1234').digest('hex');
    var now = new Date().toISOString();
    var initialData = {
      users: [
        {
          id: 1,
          name: 'Elena Rostova',
          email: 'demo@business.com',
          password_hash: demoHash,
          token: null,
          created_at: now
        }
      ],
      profiles: {
        '1': {
          firm_name: 'Velvet Bean Roastery',
          owner_name: 'Elena Rostova',
          business_type: 'Restaurant / Food',
          phone: '2065550184',
          email: 'elena@velvetbean.com',
          address: 'Pike Place Market, Seattle, WA',
          updated_at: now
        }
      },
      inputs: {
        '1': {
          fixed_costs: 12500,
          selling_price: 6.5,
          variable_cost: 2.1,
          expected_volume: 3800,
          target_profit: 5000,
          currency: '$',
          unit_label: 'cups',
          updated_at: now
        }
      }
    };
    fs.writeFileSync(DB_FILE, JSON.stringify(initialData, null, 2));
  }
}

initDb();

function readDb() {
  try {
    return JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
  } catch (e) {
    return { users: [], profiles: {}, inputs: {} };
  }
}

function writeDb(data) {
  try {
    fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2));
    return true;
  } catch (e) {
    return false;
  }
}

// Pure math calculation engine
function calculateModel(f, p, v, q, t) {
  f = Math.max(0, Number(f) || 0);
  p = Math.max(0, Number(p) || 0);
  v = Math.max(0, Number(v) || 0);
  q = Math.max(0, Number(q) || 0);
  t = Math.max(0, Number(t) || 0);

  var cm = p - v;
  var cmRatio = p > 0 ? (cm / p) : 0;
  var isValid = cm > 0;

  var beUnits = isValid ? (f / cm) : null;
  var beRevenue = isValid ? (beUnits * p) : null;

  var expRevenue = p * q;
  var expVarCost = v * q;
  var expTotalCost = f + expVarCost;
  var expProfit = expRevenue - expTotalCost;

  var mosUnits = isValid ? (q - beUnits) : null;
  var mosRatio = (isValid && q > 0) ? ((q - beUnits) / q) : 0;
  var mosRevenue = isValid ? (mosUnits * p) : null;

  var targetUnits = isValid ? ((f + t) / cm) : null;
  var targetRevenue = isValid ? (targetUnits * p) : null;

  var totalCm = q * cm;
  var dol = (expProfit > 0) ? (totalCm / expProfit) : null;

  return {
    fixed_costs: f,
    selling_price: p,
    variable_cost: v,
    expected_volume: q,
    target_profit: t,
    contribution_margin: cm,
    contribution_margin_ratio: cmRatio * 100,
    is_valid: isValid,
    break_even_units: beUnits,
    break_even_revenue: beRevenue,
    expected_revenue: expRevenue,
    expected_total_cost: expTotalCost,
    expected_profit: expProfit,
    margin_of_safety_units: mosUnits,
    margin_of_safety_ratio: mosRatio * 100,
    target_units: targetUnits,
    target_revenue: targetRevenue,
    degree_of_operating_leverage: dol
  };
}

// Helpers
function sendJson(res, status, data) {
  res.writeHead(status, {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Requested-With'
  });
  res.end(JSON.stringify(data));
}

function parseBody(req, callback) {
  var body = '';
  req.on('data', function(chunk) { body += chunk; });
  req.on('end', function() {
    try {
      callback(null, body ? JSON.parse(body) : {});
    } catch (e) {
      callback(e, {});
    }
  });
}

function getUserFromToken(req, db) {
  var auth = req.headers['authorization'] || '';
  if (auth.indexOf('Bearer ') === 0) {
    var token = auth.substring(7).trim();
    for (var i = 0; i < db.users.length; i++) {
      if (db.users[i].token === token) return db.users[i];
    }
  }
  return null;
}

var MIME_TYPES = {
  '.html': 'text/html',
  '.css': 'text/css',
  '.js': 'application/javascript',
  '.json': 'application/json',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon'
};

var server = http.createServer(function(req, res) {
  // Handle CORS Preflight
  if (req.method === 'OPTIONS') {
    res.writeHead(200, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Requested-With'
    });
    return res.end();
  }

  var parsedUrl = new URL(req.url, 'http://localhost:' + PORT);
  var pathname = parsedUrl.pathname;

  // 1. Health check
  if (pathname === '/api/health' && req.method === 'GET') {
    return sendJson(res, 200, {
      status: 'ok',
      backend: 'Node.js native HTTP',
      timestamp: new Date().toISOString(),
      database: 'database.json'
    });
  }

  // 2. Auth: Sign up
  if (pathname === '/api/auth/signup' && req.method === 'POST') {
    return parseBody(req, function(err, body) {
      var name = (body.name || '').trim();
      var email = (body.email || '').trim().toLowerCase();
      var password = body.password || '';
      var passHash = body.password_hash || (password ? crypto.createHash('sha256').update(password).digest('hex') : null);

      if (!name || !email || !passHash) {
        return sendJson(res, 400, { error: 'Name, email, and password required.' });
      }

      var db = readDb();
      for (var i = 0; i < db.users.length; i++) {
        if (db.users[i].email === email) {
          return sendJson(res, 409, { error: 'An account with this email already exists.' });
        }
      }

      var token = crypto.randomBytes(24).toString('hex');
      var now = new Date().toISOString();
      var newId = db.users.length + 1;

      var newUser = {
        id: newId,
        name: name,
        email: email,
        password_hash: passHash,
        token: token,
        created_at: now
      };
      db.users.push(newUser);

      var prof = body.profile || {};
      db.profiles[String(newId)] = {
        firm_name: prof.firm_name || (name + "'s Business"),
        owner_name: prof.owner_name || name,
        business_type: prof.business_type || 'Retail',
        phone: prof.phone || '5551234567',
        email: prof.email || email,
        address: prof.address || 'Headquarters',
        updated_at: now
      };

      var inp = body.inputs || {};
      db.inputs[String(newId)] = {
        fixed_costs: Number(inp.fixed_costs) || 12000,
        selling_price: Number(inp.selling_price) || 40,
        variable_cost: Number(inp.variable_cost) || 15,
        expected_volume: Number(inp.expected_volume) || 600,
        target_profit: Number(inp.target_profit) || 3000,
        currency: inp.currency || '$',
        unit_label: inp.unit_label || 'units',
        updated_at: now
      };

      writeDb(db);

      return sendJson(res, 201, {
        success: true,
        token: token,
        user: { id: newId, name: name, email: email },
        profile: db.profiles[String(newId)],
        inputs: db.inputs[String(newId)]
      });
    });
  }

  // 3. Auth: Login
  if (pathname === '/api/auth/login' && req.method === 'POST') {
    return parseBody(req, function(err, body) {
      var email = (body.email || '').trim().toLowerCase();
      var password = body.password || '';
      var passHash = body.password_hash || (password ? crypto.createHash('sha256').update(password).digest('hex') : null);

      var db = readDb();
      var user = null;
      for (var i = 0; i < db.users.length; i++) {
        if (db.users[i].email === email) {
          user = db.users[i];
          break;
        }
      }

      if (!user || user.password_hash !== passHash) {
        return sendJson(res, 401, { error: 'Invalid email or password.' });
      }

      var token = crypto.randomBytes(24).toString('hex');
      user.token = token;
      writeDb(db);

      return sendJson(res, 200, {
        success: true,
        token: token,
        user: { id: user.id, name: user.name, email: user.email },
        profile: db.profiles[String(user.id)] || {},
        inputs: db.inputs[String(user.id)] || {}
      });
    });
  }

  // 4. Auth: Current user (GET /api/auth/me)
  if (pathname === '/api/auth/me' && req.method === 'GET') {
    var db = readDb();
    var user = getUserFromToken(req, db);
    if (!user) {
      return sendJson(res, 401, { error: 'Unauthorized', authenticated: false });
    }
    return sendJson(res, 200, {
      authenticated: true,
      user: { id: user.id, name: user.name, email: user.email },
      profile: db.profiles[String(user.id)] || {},
      inputs: db.inputs[String(user.id)] || {}
    });
  }

  // 5. Auth: Logout
  if (pathname === '/api/auth/logout' && req.method === 'POST') {
    var db = readDb();
    var user = getUserFromToken(req, db);
    if (user) {
      user.token = null;
      writeDb(db);
    }
    return sendJson(res, 200, { success: true });
  }

  // 6. Profile GET / POST
  if (pathname === '/api/profile') {
    var db = readDb();
    var user = getUserFromToken(req, db);
    if (!user) return sendJson(res, 401, { error: 'Unauthorized' });

    if (req.method === 'GET') {
      return sendJson(res, 200, { success: true, profile: db.profiles[String(user.id)] || {} });
    }
    if (req.method === 'POST' || req.method === 'PUT') {
      return parseBody(req, function(err, body) {
        var prof = db.profiles[String(user.id)] || {};
        prof.firm_name = body.firm_name || prof.firm_name;
        prof.owner_name = body.owner_name || prof.owner_name;
        prof.business_type = body.business_type || prof.business_type;
        prof.phone = body.phone || prof.phone;
        prof.email = body.email || prof.email;
        prof.address = body.address || prof.address;
        prof.updated_at = new Date().toISOString();
        db.profiles[String(user.id)] = prof;
        writeDb(db);
        return sendJson(res, 200, { success: true, profile: prof });
      });
    }
  }

  // 7. Inputs GET / POST
  if (pathname === '/api/inputs') {
    var db = readDb();
    var user = getUserFromToken(req, db);
    if (!user) return sendJson(res, 401, { error: 'Unauthorized' });

    if (req.method === 'GET') {
      return sendJson(res, 200, { success: true, inputs: db.inputs[String(user.id)] || {} });
    }
    if (req.method === 'POST' || req.method === 'PUT') {
      return parseBody(req, function(err, body) {
        var inp = db.inputs[String(user.id)] || {};
        inp.fixed_costs = Number(body.fixed_costs) || 0;
        inp.selling_price = Number(body.selling_price) || 0;
        inp.variable_cost = Number(body.variable_cost) || 0;
        inp.expected_volume = Number(body.expected_volume) || 0;
        inp.target_profit = Number(body.target_profit) || 0;
        inp.currency = body.currency || '$';
        inp.unit_label = body.unit_label || 'units';
        inp.updated_at = new Date().toISOString();
        db.inputs[String(user.id)] = inp;
        writeDb(db);
        return sendJson(res, 200, { success: true, inputs: inp });
      });
    }
  }

  // 8. Server-side Calculation
  if (pathname === '/api/calculate' && req.method === 'POST') {
    return parseBody(req, function(err, body) {
      var calc = calculateModel(
        body.fixed_costs,
        body.selling_price,
        body.variable_cost,
        body.expected_volume,
        body.target_profit
      );
      return sendJson(res, 200, { success: true, calculation: calc });
    });
  }

  // Static File Serving
  var safePath = pathname === '/' ? 'index.html' : pathname.replace(/^\/+/, '');
  var filePath = path.join(__dirname, safePath);

  fs.stat(filePath, function(err, stats) {
    if (!err && stats.isFile()) {
      var ext = path.extname(filePath).toLowerCase();
      var contentType = MIME_TYPES[ext] || 'application/octet-stream';
      res.writeHead(200, { 'Content-Type': contentType });
      fs.createReadStream(filePath).pipe(res);
    } else {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('404 Not Found');
    }
  });
});

server.listen(PORT, function() {
  console.log('\n=======================================================');
  console.log(' Break-Even Point Predictor - Node.js Server Active');
  console.log(' Local URL: http://localhost:' + PORT + '/');
  console.log(' JSON DB:   ' + DB_FILE);
  console.log(' Press Ctrl+C to stop the server');
  console.log('=======================================================\n');
});
