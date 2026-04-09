const express = require('express');
const session = require('express-session');
const bcrypt = require('bcryptjs');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const dayjs = require('dayjs');

const app = express();
const PORT = 3000;

const baseDir = __dirname;
const viewsDir = path.join(baseDir, 'views');
const publicDir = path.join(baseDir, 'public');
const dbPath = path.join(baseDir, 'refrigerantes.db');

if (!fs.existsSync(viewsDir)) fs.mkdirSync(viewsDir, { recursive: true });
if (!fs.existsSync(publicDir)) fs.mkdirSync(publicDir, { recursive: true });

const db = new sqlite3.Database(dbPath);

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(publicDir));
app.use(
  session({
    secret: 'segredo-super-seguro-troque-isso',
    resave: false,
    saveUninitialized: false,
  })
);

app.set('view engine', 'ejs');
app.set('views', viewsDir);

const templates = {
  'login.ejs': `
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Login</title>
  <style>
    :root {
      --bg:#051426;
      --card:#0d2b4f;
      --border:rgba(255,255,255,.12);
      --text:#eef4ff;
      --muted:#b8c8e6;
      --accent:#2563eb;
      --accent2:#5a8dff;
      --danger:#ff8a97;
    }
    * { box-sizing:border-box; }
    body {
      font-family: Arial, sans-serif;
      margin:0;
      min-height:100vh;
      background: radial-gradient(circle at top, #144a85 0%, #0a2342 32%, #051426 100%);
      color:var(--text);
      display:flex;
      align-items:center;
      justify-content:center;
      padding:24px;
    }
    .box {
      width:100%;
      max-width:520px;
      background:rgba(13,43,79,.95);
      padding:32px;
      border-radius:24px;
      box-shadow:0 18px 45px rgba(0,0,0,.35);
      border:1px solid var(--border);
    }
    .brand { text-align:center; margin-bottom:22px; }
    .logo-wrap {
      width:92px; height:92px; margin:0 auto 16px; background:#fff; border-radius:22px;
      display:flex; align-items:center; justify-content:center; box-shadow:0 10px 25px rgba(0,0,0,.25);
      overflow:hidden;
    }
    .logo-wrap img { width:72px; height:72px; object-fit:contain; }
    h1 { margin:0 0 8px; font-size:30px; }
    .subtitle { color:var(--muted); line-height:1.5; }
    input, button {
      width:100%; padding:14px; margin-top:12px; border-radius:14px;
      border:1px solid var(--border); box-sizing:border-box;
      background:rgba(255,255,255,.05); color:var(--text);
    }
    input::placeholder { color:#97afd6; }
    button {
      background:linear-gradient(135deg, var(--accent), var(--accent2));
      color:#fff; border:none; cursor:pointer; font-weight:700;
    }
    .msg {
      background:rgba(255,138,151,.15); color:#ffd8dd; padding:12px; border-radius:12px;
      margin-top:12px; border:1px solid rgba(255,138,151,.18);
    }
    .link {
      margin-top:18px;
      text-align:center;
    }
    .link a {
      color:#d7e6ff;
      text-decoration:none;
      font-weight:700;
    }
  </style>
</head>
<body>
  <div class="box">
    <div class="brand">
      <div class="logo-wrap">
        <img src="/logo-empresa.jpg" alt="Logo da empresa" />
      </div>
      <h1>Controle de Bebidas</h1>
      <p class="subtitle">Login do sistema interno.</p>
    </div>

    <% if (error) { %>
      <div class="msg"><%= error %></div>
    <% } %>

    <form method="POST" action="/login">
      <input type="text" name="username" placeholder="Usuário" required />
      <input type="password" name="password" placeholder="Senha" required />
      <button type="submit">Entrar</button>
    </form>

    <div class="link">
      <a href="/register">Criar meu usuário</a>
    </div>
  </div>
</body>
</html>
`,

  'register.ejs': `
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Criar meu usuário</title>
  <style>
    :root {
      --bg:#051426;
      --card:#0d2b4f;
      --border:rgba(255,255,255,.12);
      --text:#eef4ff;
      --muted:#b8c8e6;
      --accent:#2563eb;
      --accent2:#5a8dff;
      --success:#25c18c;
      --danger:#ff8a97;
    }
    * { box-sizing:border-box; }
    body {
      font-family: Arial, sans-serif;
      background: radial-gradient(circle at top, #144a85 0%, #0a2342 32%, #051426 100%);
      margin:0;
      min-height:100vh;
      color:var(--text);
      display:flex;
      align-items:center;
      justify-content:center;
      padding:24px;
    }
    .box {
      max-width:560px;
      width:100%;
      background:rgba(13,43,79,.95);
      padding:30px;
      border-radius:24px;
      box-shadow:0 18px 45px rgba(0,0,0,.35);
      border:1px solid var(--border);
    }
    h1 { margin-top:0; font-size:28px; }
    p { color:var(--muted); }
    input, button {
      width:100%;
      padding:14px;
      margin-top:12px;
      border-radius:14px;
      border:1px solid var(--border);
      box-sizing:border-box;
      background:rgba(255,255,255,.05);
      color:var(--text);
    }
    input::placeholder { color:#97afd6; }
    button {
      background:linear-gradient(135deg, var(--accent), var(--accent2));
      color:#fff;
      border:none;
      cursor:pointer;
      font-weight:700;
    }
    a { text-decoration:none; color:#d7e6ff; }
    .msg {
      background:rgba(37,193,140,.15);
      color:#defff2;
      padding:12px;
      border-radius:12px;
      margin-top:12px;
      border:1px solid rgba(37,193,140,.18);
    }
    .error {
      background:rgba(255,138,151,.15);
      color:#ffd8dd;
      padding:12px;
      border-radius:12px;
      margin-top:12px;
      border:1px solid rgba(255,138,151,.18);
    }
  </style>
</head>
<body>
  <div class="box">
    <h1>Criar meu usuário</h1>
    <p>Preencha os dados para criar seu acesso como colaborador.</p>

    <% if (success) { %><div class="msg"><%= success %></div><% } %>
    <% if (error) { %><div class="error"><%= error %></div><% } %>

    <form method="POST" action="/register">
      <input type="text" name="name" placeholder="Nome completo" required />
      <input type="text" name="username" placeholder="Usuário" required />
      <input type="password" name="password" placeholder="Senha" required />
      <button type="submit">Criar meu usuário</button>
    </form>

    <p><a href="/">Voltar ao login</a></p>
  </div>
</body>
</html>
`,

  'dashboard.ejs': `
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Painel</title>
  <style>
    :root {
      --bg:#051426;
      --card:#0d2b4f;
      --border:rgba(255,255,255,.12);
      --text:#eef4ff;
      --muted:#b8c8e6;
      --accent:#2563eb;
      --accent2:#5a8dff;
      --danger:#ff8a97;
      --success:#25c18c;
    }
    * { box-sizing:border-box; }
    body {
      font-family: Arial, sans-serif;
      background: radial-gradient(circle at top, #144a85 0%, #0a2342 32%, #051426 100%);
      margin:0;
      color:var(--text);
    }
    .container { max-width:1250px; margin:34px auto; padding:0 16px; }
    .card {
      background:rgba(13,43,79,.95);
      padding:24px;
      border-radius:20px;
      box-shadow:0 12px 34px rgba(0,0,0,.25);
      margin-bottom:20px;
      border:1px solid var(--border);
    }
    .top {
      display:flex;
      justify-content:space-between;
      align-items:center;
      gap:12px;
      flex-wrap:wrap;
    }
    .brand {
      display:flex;
      align-items:center;
      gap:14px;
    }
    .brand-box {
      width:70px;
      height:70px;
      border-radius:18px;
      background:#fff;
      display:flex;
      align-items:center;
      justify-content:center;
      overflow:hidden;
      box-shadow:0 10px 26px rgba(0,0,0,.25);
    }
    .brand-box img {
      width:56px;
      height:56px;
      object-fit:contain;
    }
    .muted { color:var(--muted); }
    h1,h2,h3 { margin-top:0; }
    button, select, input {
      padding:12px 14px;
      border-radius:12px;
      border:1px solid var(--border);
      background:rgba(255,255,255,.05);
      color:var(--text);
    }
    button {
      background:linear-gradient(135deg, var(--accent), var(--accent2));
      color:#fff;
      border:none;
      cursor:pointer;
      font-weight:700;
    }
    .btn-soft { background:linear-gradient(135deg, #284f88, #3f6fb1); }
    .btn-danger { background:linear-gradient(135deg, #cc314f, #ff6b84); }
    .msg {
      background:rgba(37,193,140,.15);
      color:#defff2;
      padding:12px;
      border-radius:12px;
      border:1px solid rgba(37,193,140,.18);
    }
    .grid {
      display:grid;
      grid-template-columns:repeat(auto-fit, minmax(260px, 1fr));
      gap:16px;
    }
    .stats {
      display:grid;
      grid-template-columns:repeat(auto-fit, minmax(220px, 1fr));
      gap:16px;
    }
    .stat-card, .info {
      background:linear-gradient(135deg, rgba(255,255,255,.06), rgba(255,255,255,.03));
      border:1px solid var(--border);
      border-radius:16px;
      padding:18px;
    }
    .value, .price {
      font-size:28px;
      font-weight:800;
      margin-top:6px;
    }
    .pill {
      display:inline-block;
      padding:6px 10px;
      border-radius:999px;
      background:rgba(255,255,255,.08);
      color:#dce8ff;
      font-size:12px;
    }
    .product-image {
      width:100%;
      max-width:180px;
      height:180px;
      object-fit:contain;
      background:#fff;
      border-radius:14px;
      padding:10px;
      display:block;
      margin:12px 0;
    }
    .form-row {
      display:flex;
      gap:10px;
      flex-wrap:wrap;
      align-items:end;
    }
    .form-row > div {
      flex:1;
      min-width:180px;
    }
    .actions {
      display:flex;
      gap:10px;
      flex-wrap:wrap;
      align-items:center;
    }
    table {
      width:100%;
      border-collapse:collapse;
      margin-top:16px;
    }
    th, td {
      text-align:left;
      border-bottom:1px solid rgba(255,255,255,.08);
      padding:12px 8px;
      vertical-align:top;
    }
    a { text-decoration:none; }
  </style>
</head>
<body>
  <div class="container">
    <div class="card top">
      <div class="brand">
        <div class="brand-box"><img src="/logo-empresa.jpg" alt="Logo da empresa" /></div>
        <div>
          <h1>Controle de Bebidas</h1>
          <p class="muted">Olá, <strong><%= user.name %></strong> · Perfil: <strong><%= user.role %></strong></p>
        </div>
      </div>
      <div class="actions">
        <a href="/logout"><button class="btn-soft">Sair</button></a>
      </div>
    </div>

    <% if (message) { %>
      <div class="card"><div class="msg"><%= message %></div></div>
    <% } %>

    <% if (user.role === 'admin' || user.role === 'finance') { %>
      <div class="card">
        <h2>Painel de consumo por colaborador</h2>
        <div class="stats">
          <% summaryByUser.forEach(item => { %>
            <div class="stat-card">
              <span class="pill"><%= item.name %></span>
              <div class="value">R$ <%= Number(item.total_value || 0).toFixed(2).replace('.', ',') %></div>
              <div class="muted"><%= item.total_items || 0 %> retirada(s)</div>
            </div>
          <% }) %>
        </div>
      </div>
    <% } %>

    <div class="card">
      <h2>Produtos cadastrados</h2>
      <div class="grid">
        <% products.forEach(item => { %>
          <div class="info">
            <span class="pill">Produto</span>

            <% if (item.image_url) { %>
              <img src="<%= item.image_url %>" alt="<%= item.name %>" class="product-image" />
            <% } %>

            <h3 style="margin:10px 0 0;"><%= item.name %></h3>
            <div class="price">R$ <%= Number(item.price).toFixed(2).replace('.', ',') %></div>

            <% if (user.role === 'admin') { %>
              <form method="POST" action="/admin/products/<%= item.id %>/update" style="margin-top:14px;">
                <div class="form-row">
                  <div><input type="text" name="name" value="<%= item.name %>" required /></div>
                  <div><input type="number" step="0.01" min="0" name="price" value="<%= Number(item.price).toFixed(2) %>" required /></div>
                  <div><input type="text" name="image_url" value="<%= item.image_url || '' %>" placeholder="/produtos/coca-350ml.jpg" /></div>
                </div>
                <div class="actions" style="margin-top:10px;">
                  <button type="submit">Salvar</button>
                </div>
              </form>

              <form method="POST" action="/admin/products/<%= item.id %>/delete" onsubmit="return confirm('Deseja excluir este produto?');" style="margin-top:10px;">
                <button type="submit" class="btn-danger">Excluir produto</button>
              </form>
            <% } %>
          </div>
        <% }) %>
      </div>
    </div>

    <% if (user.role === 'employee') { %>
      <div class="card">
        <h2>Registrar retirada</h2>
        <form method="POST" action="/withdraw">
          <div class="form-row">
            <div>
              <label>Escolha o item</label><br /><br />
              <select name="item_id" required>
                <% products.forEach(item => { %>
                  <option value="<%= item.id %>"><%= item.name %> - R$ <%= Number(item.price).toFixed(2).replace('.', ',') %></option>
                <% }) %>
              </select>
            </div>
            <div style="max-width:220px;">
              <label>&nbsp;</label><br /><br />
              <button type="submit">Marcar retirada</button>
            </div>
          </div>
        </form>
      </div>

      <div class="card">
        <h2>Minhas retiradas</h2>
        <table>
          <thead>
            <tr>
              <th>Data</th>
              <th>Item</th>
              <th>Valor</th>
            </tr>
          </thead>
          <tbody>
            <% withdrawals.forEach(item => { %>
              <tr>
                <td><%= item.created_at %></td>
                <td><%= item.item_name %></td>
                <td>R$ <%= Number(item.item_price || 0).toFixed(2).replace('.', ',') %></td>
              </tr>
            <% }) %>
          </tbody>
        </table>
      </div>
    <% } %>

    <% if (user.role === 'admin') { %>
      <div class="card">
        <h2>Cadastrar usuário</h2>
        <form method="POST" action="/admin/users">
          <div class="form-row">
            <div>
              <label>Nome completo</label><br /><br />
              <input type="text" name="name" placeholder="Nome completo" required />
            </div>
            <div>
              <label>Usuário</label><br /><br />
              <input type="text" name="username" placeholder="usuario" required />
            </div>
            <div>
              <label>Senha</label><br /><br />
              <input type="password" name="password" placeholder="Senha" required />
            </div>
            <div>
              <label>Perfil</label><br /><br />
              <select name="role" required>
                <option value="employee">Colaborador</option>
                <option value="finance">Financeiro</option>
                <option value="admin">Admin</option>
              </select>
            </div>
            <div style="max-width:220px;">
              <label>&nbsp;</label><br /><br />
              <button type="submit">Cadastrar usuário</button>
            </div>
          </div>
        </form>
      </div>

      <div class="card">
        <h2>Usuários cadastrados</h2>
        <table>
          <thead>
            <tr>
              <th>Nome</th>
              <th>Usuário</th>
              <th>Perfil</th>
            </tr>
          </thead>
          <tbody>
            <% users.forEach(item => { %>
              <tr>
                <td><%= item.name %></td>
                <td><%= item.username %></td>
                <td><%= item.role %></td>
              </tr>
            <% }) %>
          </tbody>
        </table>
      </div>

      <div class="card">
        <h2>Cadastrar produto</h2>
        <form method="POST" action="/admin/products">
          <div class="form-row">
            <div>
              <label>Nome do produto</label><br /><br />
              <input type="text" name="name" placeholder="Ex.: Coca-Cola 350ml" required />
            </div>
            <div>
              <label>Preço</label><br /><br />
              <input type="number" step="0.01" min="0" name="price" placeholder="0.00" required />
            </div>
            <div>
              <label>URL/caminho da imagem</label><br /><br />
              <input type="text" name="image_url" placeholder="/produtos/coca-350ml.jpg" />
            </div>
            <div style="max-width:220px;">
              <label>&nbsp;</label><br /><br />
              <button type="submit">Cadastrar produto</button>
            </div>
          </div>
        </form>
      </div>
    <% } %>

    <% if (user.role === 'admin' || user.role === 'finance') { %>
      <div class="card">
        <h2>Relatórios</h2>
        <form method="GET" action="/reports/xlsx">
  <div class="form-row">
    <div>
      <label>Mês</label><br /><br />
      <input type="month" name="month" value="<%= new Date().toISOString().slice(0,7) %>" required />
    </div>
    <div style="max-width:240px;">
      <label>&nbsp;</label><br /><br />
      <button type="submit">Baixar Excel</button>
    </div>
  </div>
</form>

        <form method="GET" action="/reports/pdf" style="margin-top:14px;">
  <div class="form-row">
    <div>
      <label>Mês</label><br /><br />
      <input type="month" name="month" value="<%= new Date().toISOString().slice(0,7) %>" required />
    </div>
    <div style="max-width:240px;">
      <label>&nbsp;</label><br /><br />
      <button type="submit" class="btn-soft">Baixar PDF</button>
    </div>
  </div>
</form>
      </div>

      <div class="card">
        <h2>Lançamentos recentes</h2>
        <table>
          <thead>
            <tr>
              <th>Data</th>
              <th>Colaborador</th>
              <th>Usuário</th>
              <th>Item</th>
              <th>Valor</th>
              <% if (user.role === 'admin') { %><th>Ação</th><% } %>
            </tr>
          </thead>
          <tbody>
            <% withdrawalsAll.forEach(item => { %>
              <tr>
                <td><%= item.created_at %></td>
                <td><%= item.name %></td>
                <td><%= item.username %></td>
                <td><%= item.item_name %></td>
                <td>R$ <%= Number(item.item_price || 0).toFixed(2).replace('.', ',') %></td>
                <% if (user.role === 'admin') { %>
                  <td>
                    <form method="POST" action="/admin/withdrawals/<%= item.id %>/delete" onsubmit="return confirm('Deseja excluir este lançamento?');">
                      <button type="submit" class="btn-danger">Excluir</button>
                    </form>
                  </td>
                <% } %>
              </tr>
            <% }) %>
          </tbody>
        </table>
      </div>
    <% } %>
  </div>
</body>
</html>
`
};

for (const [fileName, content] of Object.entries(templates)) {
  fs.writeFileSync(path.join(viewsDir, fileName), content, 'utf8');
}

db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      username TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      role TEXT NOT NULL CHECK(role IN ('employee', 'finance', 'admin'))
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS products (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT UNIQUE NOT NULL,
      price REAL NOT NULL,
      image_url TEXT
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS withdrawals (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      user_id INTEGER NOT NULL,
      item_id INTEGER,
      item_name TEXT NOT NULL,
      item_price REAL NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL,
      FOREIGN KEY(user_id) REFERENCES users(id),
      FOREIGN KEY(item_id) REFERENCES products(id)
    )
  `);

  db.run(`ALTER TABLE products ADD COLUMN image_url TEXT`, () => {});
  db.run(`ALTER TABLE withdrawals ADD COLUMN item_price REAL NOT NULL DEFAULT 0`, () => {});
  db.run(`ALTER TABLE withdrawals ADD COLUMN item_id INTEGER`, () => {});

  const adminPasswordHash = bcrypt.hashSync('123456', 10);
  const financePasswordHash = bcrypt.hashSync('123456', 10);

  db.run(
    `INSERT OR IGNORE INTO users (id, name, username, password_hash, role) VALUES (1, 'Administrador', 'admin', ?, 'admin')`,
    [adminPasswordHash]
  );

  db.run(
    `INSERT OR IGNORE INTO users (id, name, username, password_hash, role) VALUES (2, 'Financeiro', 'financeiro', ?, 'finance')`,
    [financePasswordHash]
  );
});

function requireAuth(req, res, next) {
  if (!req.session.user) return res.redirect('/');
  next();
}

function requireAdmin(req, res, next) {
  if (!req.session.user || req.session.user.role !== 'admin') {
    return res.status(403).send('Acesso restrito ao administrador.');
  }
  next();
}

function requireFinanceOrAdmin(req, res, next) {
  if (!req.session.user || !['finance', 'admin'].includes(req.session.user.role)) {
    return res.status(403).send('Acesso restrito.');
  }
  next();
}

app.get('/', (req, res) => {
  res.render('login', { error: null });
});

app.post('/login', (req, res) => {
  const { username, password } = req.body;

  db.get(`SELECT * FROM users WHERE username = ?`, [username], (err, user) => {
    if (err) {
      return res.render('login', { error: 'Erro interno ao tentar entrar.' });
    }

    if (!user || !bcrypt.compareSync(password, user.password_hash)) {
      return res.render('login', { error: 'Usuário ou senha inválidos.' });
    }

    req.session.user = {
      id: user.id,
      name: user.name,
      username: user.username,
      role: user.role,
    };

    res.redirect('/dashboard');
  });
});

app.get('/register', (req, res) => {
  res.render('register', { success: null, error: null });
});

app.post('/register', (req, res) => {
  const { name, username, password } = req.body;

  if (!name || !username || !password) {
    return res.render('register', {
      success: null,
      error: 'Preencha todos os campos.'
    });
  }

  const passwordHash = bcrypt.hashSync(password, 10);

  db.run(
    `INSERT INTO users (name, username, password_hash, role) VALUES (?, ?, ?, 'employee')`,
    [name.trim(), username.trim(), passwordHash],
    function (err) {
      if (err) {
        return res.render('register', {
          success: null,
          error: 'Não foi possível criar o usuário. Esse login pode já existir.'
        });
      }

      res.render('register', {
        success: 'Usuário criado com sucesso. Agora você já pode entrar no sistema.',
        error: null
      });
    }
  );
});

app.get('/dashboard', requireAuth, (req, res) => {
  const user = req.session.user;
  const message = req.session.message || null;
  req.session.message = null;

  db.all(`SELECT id, name, price, image_url FROM products ORDER BY name ASC`, [], (productErr, products) => {
    if (productErr) products = [];

    if (user.role === 'employee') {
      db.all(
        `SELECT id, item_name, item_price, created_at FROM withdrawals WHERE user_id = ? ORDER BY datetime(created_at) DESC LIMIT 20`,
        [user.id],
        (err, withdrawals) => {
          if (err) withdrawals = [];
          res.render('dashboard', {
            user,
            products,
            withdrawals,
            withdrawalsAll: [],
            summaryByUser: [],
            users: [],
            message
          });
        }
      );
      return;
    }

    db.all(
      `
      SELECT w.id, w.created_at, w.item_name, w.item_price, u.name, u.username
      FROM withdrawals w
      INNER JOIN users u ON u.id = w.user_id
      ORDER BY datetime(w.created_at) DESC
      LIMIT 100
      `,
      [],
      (withdrawErr, withdrawalsAll) => {
        if (withdrawErr) withdrawalsAll = [];

        db.all(
          `
          SELECT
            u.name,
            COUNT(w.id) AS total_items,
            COALESCE(SUM(w.item_price), 0) AS total_value
          FROM users u
          LEFT JOIN withdrawals w ON w.user_id = u.id
          WHERE u.role = 'employee'
          GROUP BY u.id, u.name
          ORDER BY total_value DESC, total_items DESC, u.name ASC
          `,
          [],
          (summaryErr, summaryByUser) => {
            if (summaryErr) summaryByUser = [];

            const finishRender = (usersList) => {
              res.render('dashboard', {
                user,
                products,
                withdrawals: [],
                withdrawalsAll,
                summaryByUser,
                users: usersList || [],
                message
              });
            };

            if (user.role === 'admin') {
              db.all(
                `SELECT id, name, username, role FROM users ORDER BY role ASC, name ASC`,
                [],
                (usersErr, usersList) => {
                  if (usersErr) usersList = [];
                  finishRender(usersList);
                }
              );
            } else {
              finishRender([]);
            }
          }
        );
      }
    );
  });
});

app.post('/withdraw', requireAuth, (req, res) => {
  const user = req.session.user;

  if (user.role !== 'employee') {
    return res.status(403).send('Somente colaboradores podem registrar retiradas.');
  }

  const { item_id } = req.body;
  const createdAt = dayjs().format('YYYY-MM-DD HH:mm:ss');

  db.get(`SELECT * FROM products WHERE id = ?`, [item_id], (productErr, product) => {
    if (productErr || !product) {
      req.session.message = 'Produto não encontrado.';
      return res.redirect('/dashboard');
    }

    db.run(
      `INSERT INTO withdrawals (user_id, item_id, item_name, item_price, created_at) VALUES (?, ?, ?, ?, ?)`,
      [user.id, product.id, product.name, product.price, createdAt],
      function (err) {
        if (err) {
          req.session.message = 'Erro ao registrar retirada.';
          return res.redirect('/dashboard');
        }

        req.session.message = `Retirada registrada com sucesso: ${product.name} - R$ ${Number(product.price).toFixed(2).replace('.', ',')}.`;
        res.redirect('/dashboard');
      }
    );
  });
});

app.post('/admin/users', requireAdmin, (req, res) => {
  const { name, username, password, role } = req.body;

  if (!name || !username || !password || !role) {
    req.session.message = 'Preencha todos os campos do usuário.';
    return res.redirect('/dashboard');
  }

  if (!['admin', 'finance', 'employee'].includes(role)) {
    req.session.message = 'Perfil inválido.';
    return res.redirect('/dashboard');
  }

  const passwordHash = bcrypt.hashSync(password, 10);

  db.run(
    `INSERT INTO users (name, username, password_hash, role) VALUES (?, ?, ?, ?)`,
    [name.trim(), username.trim(), passwordHash, role],
    function (err) {
      if (err) {
        req.session.message = 'Não foi possível cadastrar o usuário.';
        return res.redirect('/dashboard');
      }

      req.session.message = 'Usuário cadastrado com sucesso.';
      res.redirect('/dashboard');
    }
  );
});

app.post('/admin/products', requireAdmin, (req, res) => {
  const { name, price, image_url } = req.body;

  if (!name || price === undefined || price === '') {
    req.session.message = 'Preencha nome e preço do produto.';
    return res.redirect('/dashboard');
  }

  db.run(
    `INSERT INTO products (name, price, image_url) VALUES (?, ?, ?)`,
    [name.trim(), Number(price), image_url ? image_url.trim() : ''],
    function (err) {
      if (err) {
        req.session.message = 'Não foi possível cadastrar o produto.';
        return res.redirect('/dashboard');
      }

      req.session.message = 'Produto cadastrado com sucesso.';
      res.redirect('/dashboard');
    }
  );
});

app.post('/admin/products/:id/update', requireAdmin, (req, res) => {
  const { id } = req.params;
  const { name, price, image_url } = req.body;

  if (!name || price === undefined || price === '') {
    req.session.message = 'Preencha nome e preço para atualizar.';
    return res.redirect('/dashboard');
  }

  db.run(
    `UPDATE products SET name = ?, price = ?, image_url = ? WHERE id = ?`,
    [name.trim(), Number(price), image_url ? image_url.trim() : '', id],
    function (err) {
      if (err) {
        req.session.message = 'Erro ao atualizar produto.';
        return res.redirect('/dashboard');
      }

      req.session.message = 'Produto atualizado com sucesso.';
      res.redirect('/dashboard');
    }
  );
});

app.post('/admin/products/:id/delete', requireAdmin, (req, res) => {
  const { id } = req.params;

  db.get(`SELECT COUNT(*) AS total FROM withdrawals WHERE item_id = ?`, [id], (countErr, row) => {
    if (countErr) {
      req.session.message = 'Erro ao validar exclusão do produto.';
      return res.redirect('/dashboard');
    }

    if (row && row.total > 0) {
      req.session.message = 'Não é possível excluir um produto que já possui lançamentos.';
      return res.redirect('/dashboard');
    }

    db.run(`DELETE FROM products WHERE id = ?`, [id], function (err) {
      if (err) {
        req.session.message = 'Erro ao excluir produto.';
        return res.redirect('/dashboard');
      }

      req.session.message = 'Produto excluído com sucesso.';
      res.redirect('/dashboard');
    });
  });
});

app.post('/admin/withdrawals/:id/delete', requireAdmin, (req, res) => {
  const { id } = req.params;

  db.run(`DELETE FROM withdrawals WHERE id = ?`, [id], function (err) {
    if (err) {
      req.session.message = 'Erro ao excluir lançamento.';
      return res.redirect('/dashboard');
    }

    req.session.message = 'Lançamento excluído com sucesso.';
    res.redirect('/dashboard');
  });
});

app.get('/reports/xlsx', requireFinanceOrAdmin, (req, res) => {
  const { month } = req.query;

  if (!month || !/^\\d{4}-\\d{2}$/.test(month)) {
    return res.status(400).send('Informe o mês no formato YYYY-MM.');
  }

  const start = `${month}-01 00:00:00`;
  const end = dayjs(`${month}-01`).add(1, 'month').format('YYYY-MM-DD 00:00:00');

  db.all(
    `
    SELECT
      u.name AS Colaborador,
      u.username AS Usuario,
      w.item_name AS Item,
      w.item_price AS Valor,
      w.created_at AS DataHora
    FROM withdrawals w
    INNER JOIN users u ON u.id = w.user_id
    WHERE w.created_at >= ? AND w.created_at < ?
    ORDER BY u.name ASC, w.created_at ASC
    `,
    [start, end],
    (err, rows) => {
      if (err) return res.status(500).send('Erro ao gerar relatório.');

      const resumoPorPessoa = {};
      for (const row of rows) {
        if (!resumoPorPessoa[row.Colaborador]) {
          resumoPorPessoa[row.Colaborador] = { total: 0, valor: 0 };
        }
        resumoPorPessoa[row.Colaborador].total += 1;
        resumoPorPessoa[row.Colaborador].valor += Number(row.Valor || 0);
      }

      const resumoSheet = Object.entries(resumoPorPessoa).map(([colaborador, dados]) => ({
        Colaborador: colaborador,
        TotalRetiradas: dados.total,
        TotalEmReais: Number(dados.valor.toFixed(2)),
      }));

      const workbook = XLSX.utils.book_new();
      const detalheWs = XLSX.utils.json_to_sheet(rows);
      const resumoWs = XLSX.utils.json_to_sheet(resumoSheet);

      XLSX.utils.book_append_sheet(workbook, detalheWs, 'Detalhado');
      XLSX.utils.book_append_sheet(workbook, resumoWs, 'Resumo');

      const fileName = `relatorio-bebidas-${month}.xlsx`;
      const filePath = path.join(baseDir, fileName);

      XLSX.writeFile(workbook, filePath);

      res.download(filePath, fileName, (downloadErr) => {
        if (downloadErr) console.error(downloadErr);
        if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
      });
    }
  );
});

app.get('/reports/pdf', requireFinanceOrAdmin, (req, res) => {
  const { month } = req.query;

  if (!month || !/^\\d{4}-\\d{2}$/.test(month)) {
    return res.status(400).send('Informe o mês no formato YYYY-MM.');
  }

  const start = `${month}-01 00:00:00`;
  const end = dayjs(`${month}-01`).add(1, 'month').format('YYYY-MM-DD 00:00:00');

  db.all(
    `
    SELECT
      u.name AS Colaborador,
      u.username AS Usuario,
      w.item_name AS Item,
      w.item_price AS Valor,
      w.created_at AS DataHora
    FROM withdrawals w
    INNER JOIN users u ON u.id = w.user_id
    WHERE w.created_at >= ? AND w.created_at < ?
    ORDER BY u.name ASC, w.created_at ASC
    `,
    [start, end],
    (err, rows) => {
      if (err) return res.status(500).send('Erro ao gerar PDF.');

      const fileName = `relatorio-bebidas-${month}.pdf`;
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);

      const doc = new PDFDocument({ margin: 40, size: 'A4' });
      doc.pipe(res);

      const logoPath = path.join(publicDir, 'logo-empresa.jpg');
      if (fs.existsSync(logoPath)) {
        doc.image(logoPath, 40, 30, { fit: [60, 60] });
      }

      doc.fontSize(18).text('Relatório Mensal de Bebidas', 120, 40);
      doc.fontSize(12).text(`Mês: ${month}`, 120, 65);
      doc.moveDown(3);

      let totalGeral = 0;

      rows.forEach((row, index) => {
        totalGeral += Number(row.Valor || 0);
        doc.fontSize(10).text(
          `${index + 1}. ${row.DataHora} | ${row.Colaborador} | ${row.Usuario} | ${row.Item} | R$ ${Number(row.Valor).toFixed(2).replace('.', ',')}`
        );
      });

      doc.moveDown();
      doc.fontSize(12).text(`Total de lançamentos: ${rows.length}`);
      doc.fontSize(12).text(`Total em reais: R$ ${totalGeral.toFixed(2).replace('.', ',')}`);
      doc.end();
    }
  );
});

app.get('/logout', (req, res) => {
  req.session.destroy(() => {
    res.redirect('/');
  });
});

app.listen(PORT, () => {
  console.log(`Servidor rodando em http://localhost:${PORT}`);
});
