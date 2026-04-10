console.log('APP INICIANDO...');
const express = require('express');
const session = require('express-session');
const bcrypt = require('bcryptjs');
const { Pool } = require('pg');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const multer = require('multer');
const dayjs = require('dayjs');
const utc = require('dayjs/plugin/utc');
const timezone = require('dayjs/plugin/timezone');
dayjs.extend(utc);
dayjs.extend(timezone);
dayjs.tz.setDefault('America/Cuiaba');

const app = express();
const PORT = process.env.PORT || 3000;

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false,
});

const baseDir = __dirname;
const viewsDir = path.join(baseDir, 'views');
const publicDir = path.join(baseDir, 'public');

if (!fs.existsSync(viewsDir)) fs.mkdirSync(viewsDir, { recursive: true });
if (!fs.existsSync(publicDir)) fs.mkdirSync(publicDir, { recursive: true });

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(publicDir));

app.use(
  session({
    secret: process.env.SESSION_SECRET || 'segredo-super-seguro-troque-isso',
    resave: false,
    saveUninitialized: false,
    cookie: {
      secure: false,
      httpOnly: true,
      maxAge: 1000 * 60 * 60 * 12,
    },
  })
);

app.set('view engine', 'ejs');
app.set('views', viewsDir);

// Multer para upload de notas fiscais (armazena em mem\u00f3ria)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB m\u00e1ximo
  fileFilter: (req, file, cb) => {
    const allowed = ['application/pdf', 'image/jpeg', 'image/png', 'image/webp'];
    if (allowed.includes(file.mimetype)) cb(null, true);
    else cb(new Error('Formato n\u00e3o permitido. Envie PDF, JPG, PNG ou WEBP.'));
  }
});

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
      <input type="text" name="cpf" id="cpf-register" placeholder="CPF (000.000.000-00)" maxlength="14" required />
      <input type="text" name="username" placeholder="Usuário" required />
      <input type="password" name="password" placeholder="Senha" required />
      <button type="submit">Criar meu usuário</button>
    </form>

    <script>
      document.getElementById('cpf-register').addEventListener('input', function(e) {
        let v = e.target.value.replace(/\D/g, '').slice(0, 11);
        if (v.length > 9) v = v.replace(/(\d{3})(\d{3})(\d{3})(\d{1,2})/, '$1.$2.$3-$4');
        else if (v.length > 6) v = v.replace(/(\d{3})(\d{3})(\d{1,3})/, '$1.$2.$3');
        else if (v.length > 3) v = v.replace(/(\d{3})(\d{1,3})/, '$1.$2');
        e.target.value = v;
      });
    </script>

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
      --successBg:rgba(37,193,140,.15);
      --warningBg:rgba(255,193,7,.16);
      --warningBorder:rgba(255,193,7,.28);
      --dangerBg:rgba(255,107,132,.14);
      --dangerBorder:rgba(255,107,132,.22);
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
    select option { color:#000; }
    button {
      background:linear-gradient(135deg, var(--accent), var(--accent2));
      color:#fff;
      border:none;
      cursor:pointer;
      font-weight:700;
    }
    .btn-soft { background:linear-gradient(135deg, #284f88, #3f6fb1); }
    .btn-danger { background:linear-gradient(135deg, #cc314f, #ff6b84); }
    .btn-pix { background:linear-gradient(135deg, #0d9b73, #21c58b); }
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
    .stock-box {
      margin-top:12px;
      padding:10px 14px;
      border-radius:12px;
      font-weight:700;
    }
    .stock-ok {
      background:var(--successBg);
      border:1px solid rgba(37,193,140,.25);
      color:#defff2;
    }
    .stock-low {
      background:var(--warningBg);
      border:1px solid var(--warningBorder);
      color:#ffe9a6;
    }
    .stock-zero {
      background:var(--dangerBg);
      border:1px solid var(--dangerBorder);
      color:#ffd8dd;
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
    a { text-decoration:none; color:inherit; }
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

            <% if (user.role === 'admin' || user.role === 'finance') { %>
              <% const estoque = Number(item.stock_quantity || 0); %>
              <div class="stock-box <%= estoque <= 0 ? 'stock-zero' : estoque <= 5 ? 'stock-low' : 'stock-ok' %>">
                Estoque atual: <%= estoque %>
              </div>
            <% } %>

            <% if (user.role === 'admin') { %>
              <form method="POST" action="/admin/products/<%= item.id %>/update" style="margin-top:14px;">
                <div class="form-row">
                  <div><input type="text" name="name" value="<%= item.name %>" required /></div>
                  <div><input type="number" step="0.01" min="0" name="price" value="<%= Number(item.price).toFixed(2) %>" required /></div>
                  <div><input type="text" name="image_url" value="<%= item.image_url || '' %>" placeholder="/produtos/coca-350ml.jpg" /></div>
                  <div><input type="number" min="0" name="stock_quantity" value="<%= item.stock_quantity || 0 %>" placeholder="Estoque" /></div>
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
                  <option value="<%= item.id %>" <%= Number(item.stock_quantity || 0) <= 0 ? 'disabled' : '' %>>
                    <%= item.name %> - R$ <%= Number(item.price).toFixed(2).replace('.', ',') %>
                    <%= Number(item.stock_quantity || 0) <= 0 ? ' | INDISPON\u00cdVEL' : '' %>
                  </option>
                <% }) %>
              </select>
            </div>

            <div style="max-width:220px;">
              <label>&nbsp;</label><br /><br />
              <button type="submit">Marcar retirada</button>
            </div>

            <div style="max-width:220px;">
              <label>&nbsp;</label><br /><br />
              <button type="button" class="btn-pix" onclick="alert('Pagamento via Pix: configurar chave Pix da empresa aqui.')">
                Pagar com Pix
              </button>
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
                <td><%= dayjs(item.created_at).tz('America/Cuiaba').format('DD/MM/YYYY HH:mm:ss') %></td>
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
            <div><label>Nome completo</label><br /><br /><input type="text" name="name" placeholder="Nome completo" required /></div>
            <div><label>CPF</label><br /><br /><input type="text" name="cpf" id="cpf-admin" placeholder="000.000.000-00" maxlength="14" required /></div>
            <div><label>Usuário</label><br /><br /><input type="text" name="username" placeholder="usuario" required /></div>
            <div><label>Senha</label><br /><br /><input type="password" name="password" placeholder="Senha" required /></div>
            <div>
              <label>Perfil</label><br /><br />
              <select name="role" required>
                <option value="employee">Colaborador</option>
                <option value="finance">Financeiro</option>
                <option value="admin">Admin</option>
              </select>
            </div>
            <div style="max-width:220px;"><label>&nbsp;</label><br /><br /><button type="submit">Cadastrar usuário</button></div>
          </div>
        </form>
      </div>

      <div class="card">
        <h2>Usuários cadastrados</h2>
        <table>
          <thead>
            <tr>
              <th>Nome</th>
              <th>CPF</th>
              <th>Usu\u00e1rio</th>
              <th>Perfil</th>
              <th>A\u00e7\u00e3o</th>
            </tr>
          </thead>
          <tbody>
            <% users.forEach(item => { %>
              <tr>
                <td><%= item.name %></td>
                <td><%= item.cpf || '-' %></td>
                <td><%= item.username %></td>
                <td><%= item.role %></td>
                <td>
                  <% if (item.username !== 'admin' && item.username !== 'financeiro') { %>
                    <form method="POST" action="/admin/users/<%= item.id %>/delete" onsubmit="return confirm('Deseja realmente excluir o usu\u00e1rio ' + '<%= item.name %>' + '? Todos os lan\u00e7amentos dele ser\u00e3o removidos.')" style="margin:0;">
                      <button type="submit" class="btn-danger" style="padding:6px 12px; font-size:12px;">Excluir</button>
                    </form>
                  <% } else { %>
                    -
                  <% } %>
                </td>
              </tr>
            <% }) %>
          </tbody>
        </table>
      </div>

      <div class="card">
        <h2>Cadastrar produto</h2>
        <form method="POST" action="/admin/products">
          <div class="form-row">
            <div><label>Nome do produto</label><br /><br /><input type="text" name="name" placeholder="Ex.: Coca-Cola 350ml" required /></div>
            <div><label>Preço</label><br /><br /><input type="number" step="0.01" min="0" name="price" placeholder="0.00" required /></div>
            <div><label>URL/caminho da imagem</label><br /><br /><input type="text" name="image_url" placeholder="/produtos/coca-350ml.jpg" /></div>
            <div><label>Estoque inicial</label><br /><br /><input type="number" min="0" name="stock_quantity" placeholder="0" /></div>
            <div style="max-width:220px;"><label>&nbsp;</label><br /><br /><button type="submit">Cadastrar produto</button></div>
          </div>
        </form>
      </div>
    <% } %>

    <% if (user.role === 'admin' || user.role === 'finance') { %>
      <div class="card">
        <h2>Entrada de estoque</h2>
        <form method="POST" action="/admin/stock/add">
          <div class="form-row">
            <div>
              <label>Produto</label><br /><br />
              <select name="product_id" required>
                <% products.forEach(item => { %>
                  <option value="<%= item.id %>"><%= item.name %> | Estoque atual: <%= item.stock_quantity %></option>
                <% }) %>
              </select>
            </div>
            <div>
              <label>Quantidade recebida</label><br /><br />
              <input type="number" min="1" name="quantity" placeholder="0" required />
            </div>
            <div>
              <label>Custo total da entrega (R$)</label><br /><br />
              <input type="number" step="0.01" min="0" name="total_cost" placeholder="Ex: 150.00" required />
            </div>
            <div style="max-width:220px;">
              <label>&nbsp;</label><br /><br />
              <button type="submit">Adicionar estoque</button>
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
              <label>Data início</label><br /><br />
              <input type="date" name="start" value="<%= new Date(new Date().getFullYear(), new Date().getMonth(), 1).toISOString().slice(0,10) %>" required />
            </div>
            <div>
              <label>Data fim</label><br /><br />
              <input type="date" name="end" value="<%= new Date().toISOString().slice(0,10) %>" required />
            </div>
            <div style="max-width:240px;">
              <label>&nbsp;</label><br /><br />
              <button type="submit">Baixar Excel</button>
            </div>
          </div>
        </form>
      </div>

      <div class="card">
        <h2>Notas Fiscais</h2>

        <% if (user.role === 'admin' || user.role === 'finance') { %>
          <form method="POST" action="/invoices/upload" enctype="multipart/form-data" style="margin-bottom:20px;">
            <div class="form-row">
              <div>
                <label>M\u00eas de refer\u00eancia</label><br /><br />
                <input type="month" name="reference_month" required />
              </div>
              <div>
                <label>Arquivo (PDF, JPG, PNG)</label><br /><br />
                <input type="file" name="invoice_file" accept=".pdf,.jpg,.jpeg,.png,.webp" required style="background:rgba(255,255,255,.05); padding:10px; border-radius:12px; border:1px solid rgba(255,255,255,.12); color:#eef4ff;" />
              </div>
              <div style="max-width:240px;">
                <label>&nbsp;</label><br /><br />
                <button type="submit">Enviar nota</button>
              </div>
            </div>
          </form>
        <% } %>

        <div style="margin-bottom:14px;">
          <form method="GET" action="/dashboard" style="display:inline-flex; gap:10px; align-items:end;">
            <div>
              <label>Filtrar por m\u00eas</label><br /><br />
              <input type="month" name="invoice_month" value="<%= typeof invoiceMonth !== 'undefined' ? invoiceMonth : '' %>" />
            </div>
            <div><button type="submit">Filtrar</button></div>
          </form>
        </div>

        <% if (typeof invoices !== 'undefined' && invoices.length > 0) { %>
          <table>
            <thead>
              <tr>
                <th>M\u00eas Ref.</th>
                <th>Arquivo</th>
                <th>Enviado em</th>
                <th>Visualizar</th>
                <th>Baixar</th>
                <% if (user.role === 'admin' || user.role === 'finance') { %><th>Excluir</th><% } %>
              </tr>
            </thead>
            <tbody>
              <% invoices.forEach(inv => { %>
                <tr>
                  <td><%= inv.reference_month %></td>
                  <td><%= inv.file_name %></td>
                  <td><%= dayjs(inv.created_at).tz('America/Cuiaba').format('DD/MM/YYYY HH:mm') %></td>
                  <td><a href="/invoices/<%= inv.id %>/view" target="_blank" style="color:#5a8dff; text-decoration:none; font-weight:700;">Abrir</a></td>
                  <td><a href="/invoices/<%= inv.id %>/download" style="color:#25c18c; text-decoration:none; font-weight:700;">Baixar</a></td>
                  <% if (user.role === 'admin' || user.role === 'finance') { %>
                    <td>
                      <form method="POST" action="/invoices/<%= inv.id %>/delete" onsubmit="return confirm('Deseja excluir esta nota fiscal?')" style="margin:0;">
                        <button type="submit" class="btn-danger" style="padding:6px 12px; font-size:12px;">Excluir</button>
                      </form>
                    </td>
                  <% } %>
                </tr>
              <% }) %>
            </tbody>
          </table>
        <% } else { %>
          <p class="muted">Nenhuma nota fiscal encontrada<%= typeof invoiceMonth !== 'undefined' && invoiceMonth ? ' para o m\u00eas selecionado' : '' %>.</p>
        <% } %>
      </div>

      <div class="card">
        <h2>Lan\u00e7amentos recentes</h2>
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
                <td><%= dayjs(item.created_at).tz('America/Cuiaba').format('DD/MM/YYYY HH:mm:ss') %></td>
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
  <script>
    var cpfField = document.getElementById('cpf-admin');
    if (cpfField) {
      cpfField.addEventListener('input', function(e) {
        let v = e.target.value.replace(/\D/g, '').slice(0, 11);
        if (v.length > 9) v = v.replace(/(\d{3})(\d{3})(\d{3})(\d{1,2})/, '$1.$2.$3-$4');
        else if (v.length > 6) v = v.replace(/(\d{3})(\d{3})(\d{1,3})/, '$1.$2.$3');
        else if (v.length > 3) v = v.replace(/(\d{3})(\d{1,3})/, '$1.$2');
        e.target.value = v;
      });
    }
  </script>
</body>
</html>
`
};

for (const [fileName, content] of Object.entries(templates)) {
  fs.writeFileSync(path.join(viewsDir, fileName), content, 'utf8');
}

async function initDB() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS users (
      id SERIAL PRIMARY KEY,
      name TEXT NOT NULL,
      username TEXT UNIQUE NOT NULL,
      cpf TEXT UNIQUE,
      password_hash TEXT NOT NULL,
      role TEXT NOT NULL CHECK (role IN ('employee', 'finance', 'admin'))
    )
  `);

  // Adicionar coluna cpf caso a tabela já exista sem ela
  await pool.query(`
    DO $$
    BEGIN
      IF NOT EXISTS (
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'users' AND column_name = 'cpf'
      ) THEN
        ALTER TABLE users ADD COLUMN cpf TEXT UNIQUE;
      END IF;
    END
    $$;
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS products (
      id SERIAL PRIMARY KEY,
      name TEXT UNIQUE NOT NULL,
      price NUMERIC(10,2) NOT NULL,
      image_url TEXT,
      stock_quantity INTEGER NOT NULL DEFAULT 0
    )
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS withdrawals (
      id SERIAL PRIMARY KEY,
      user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
      item_id INTEGER REFERENCES products(id) ON DELETE SET NULL,
      item_name TEXT NOT NULL,
      item_price NUMERIC(10,2) NOT NULL DEFAULT 0,
      created_at TIMESTAMPTZ NOT NULL DEFAULT (NOW() AT TIME ZONE 'America/Cuiaba')
    )
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS stock_entries (
      id SERIAL PRIMARY KEY,
      product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
      product_name TEXT NOT NULL,
      quantity INTEGER NOT NULL,
      unit_cost NUMERIC(10,2) NOT NULL DEFAULT 0,
      total_cost NUMERIC(10,2) NOT NULL DEFAULT 0,
      created_at TIMESTAMPTZ NOT NULL DEFAULT (NOW() AT TIME ZONE 'America/Cuiaba')
    )
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS invoices (
      id SERIAL PRIMARY KEY,
      reference_month TEXT NOT NULL,
      file_name TEXT NOT NULL,
      file_type TEXT NOT NULL,
      file_data BYTEA NOT NULL,
      uploaded_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT (NOW() AT TIME ZONE 'America/Cuiaba')
    )
  `);

  const adminPasswordHash = bcrypt.hashSync('123456', 10);
  const financePasswordHash = bcrypt.hashSync('123456', 10);

  await pool.query(
    `
    INSERT INTO users (name, username, password_hash, role)
    VALUES ('Administrador', 'admin', $1, 'admin')
    ON CONFLICT (username) DO NOTHING
    `,
    [adminPasswordHash]
  );

  await pool.query(
    `
    INSERT INTO users (name, username, password_hash, role)
    VALUES ('Financeiro', 'financeiro', $1, 'finance')
    ON CONFLICT (username) DO NOTHING
    `,
    [financePasswordHash]
  );

  await pool.query(
    `
    UPDATE users
    SET password_hash = $1, role = 'admin', name = 'Administrador'
    WHERE username = 'admin'
    `,
    [adminPasswordHash]
  );

  await pool.query(
    `
    UPDATE users
    SET password_hash = $1, role = 'finance', name = 'Financeiro'
    WHERE username = 'financeiro'
    `,
    [financePasswordHash]
  );
}

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

app.post('/login', async (req, res) => {
  const { username, password } = req.body;

  try {
    const result = await pool.query(
      `SELECT * FROM users WHERE username = $1`,
      [username]
    );

    const user = result.rows[0];

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
  } catch (err) {
    console.error(err);
    res.render('login', { error: 'Erro interno ao tentar entrar.' });
  }
});

app.get('/register', (req, res) => {
  res.render('register', { success: null, error: null });
});

app.post('/register', async (req, res) => {
  const { name, cpf, username, password } = req.body;

  if (!name || !cpf || !username || !password) {
    return res.render('register', {
      success: null,
      error: 'Preencha todos os campos, incluindo o CPF.'
    });
  }

  // Limpar CPF (só dígitos)
  const cpfLimpo = cpf.replace(/\D/g, '');

  if (cpfLimpo.length !== 11) {
    return res.render('register', {
      success: null,
      error: 'CPF inválido. Informe os 11 dígitos.'
    });
  }

  // Validar CPF (dígitos verificadores)
  function validarCPF(c) {
    if (/^(\d)\1{10}$/.test(c)) return false;
    let soma = 0;
    for (let i = 0; i < 9; i++) soma += Number(c[i]) * (10 - i);
    let resto = (soma * 10) % 11;
    if (resto === 10) resto = 0;
    if (resto !== Number(c[9])) return false;
    soma = 0;
    for (let i = 0; i < 10; i++) soma += Number(c[i]) * (11 - i);
    resto = (soma * 10) % 11;
    if (resto === 10) resto = 0;
    return resto === Number(c[10]);
  }

  if (!validarCPF(cpfLimpo)) {
    return res.render('register', {
      success: null,
      error: 'CPF inválido. Verifique os dígitos informados.'
    });
  }

  // Formatar CPF para salvar padronizado
  const cpfFormatado = cpfLimpo.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');

  try {
    // Verificar se CPF já existe
    const cpfExiste = await pool.query('SELECT id FROM users WHERE cpf = $1', [cpfFormatado]);
    if (cpfExiste.rows.length > 0) {
      return res.render('register', {
        success: null,
        error: 'Já existe um usuário cadastrado com este CPF.'
      });
    }

    const passwordHash = bcrypt.hashSync(password, 10);

    await pool.query(
      `INSERT INTO users (name, cpf, username, password_hash, role)
       VALUES ($1, $2, $3, $4, 'employee')`,
      [name.trim(), cpfFormatado, username.trim(), passwordHash]
    );

    res.render('register', {
      success: 'Usuário criado com sucesso. Agora você já pode entrar no sistema.',
      error: null
    });
  } catch (err) {
    console.error(err);
    if (err.code === '23505' && err.constraint && err.constraint.includes('cpf')) {
      return res.render('register', {
        success: null,
        error: 'Já existe um usuário cadastrado com este CPF.'
      });
    }
    res.render('register', {
      success: null,
      error: 'Não foi possível criar o usuário. Esse login pode já existir.'
    });
  }
});

app.get('/dashboard', requireAuth, async (req, res) => {
  const user = req.session.user;
  const message = req.session.message || null;
  req.session.message = null;

  try {
    const productsResult = await pool.query(`
      SELECT id, name, price, image_url, stock_quantity
      FROM products
      ORDER BY name ASC
    `);

    const products = productsResult.rows;

    if (user.role === 'employee') {
      const withdrawalsResult = await pool.query(
        `SELECT id, item_name, item_price, created_at
         FROM withdrawals
         WHERE user_id = $1
         ORDER BY created_at DESC
         LIMIT 20`,
        [user.id]
      );

      return res.render('dashboard', {
        user,
        products,
        withdrawals: withdrawalsResult.rows,
        withdrawalsAll: [],
        summaryByUser: [],
        users: [],
        message,
        dayjs
      });
    }

    // Buscar notas fiscais (filtro opcional por m\u00eas)
    const invoiceMonth = req.query.invoice_month || '';
    let invoicesResult;
    if (invoiceMonth) {
      invoicesResult = await pool.query(
        'SELECT id, reference_month, file_name, file_type, created_at FROM invoices WHERE reference_month = $1 ORDER BY created_at DESC',
        [invoiceMonth]
      );
    } else {
      invoicesResult = await pool.query(
        'SELECT id, reference_month, file_name, file_type, created_at FROM invoices ORDER BY reference_month DESC, created_at DESC LIMIT 50'
      );
    }

    const withdrawalsAllResult = await pool.query(`
      SELECT w.id, w.created_at, w.item_name, w.item_price, u.name, u.username
      FROM withdrawals w
      INNER JOIN users u ON u.id = w.user_id
      ORDER BY w.created_at DESC
      LIMIT 100
    `);

    const summaryByUserResult = await pool.query(`
      SELECT
        u.name,
        COUNT(w.id) AS total_items,
        COALESCE(SUM(w.item_price), 0) AS total_value
      FROM users u
      LEFT JOIN withdrawals w ON w.user_id = u.id
      WHERE u.role = 'employee'
      GROUP BY u.id, u.name
      ORDER BY total_value DESC, total_items DESC, u.name ASC
    `);

    let users = [];
    if (user.role === 'admin') {
      const usersResult = await pool.query(`
        SELECT id, name, username, cpf, role
        FROM users
        ORDER BY role ASC, name ASC
      `);
      users = usersResult.rows;
    }

    return res.render('dashboard', {
      user,
      products,
      withdrawals: [],
      withdrawalsAll: withdrawalsAllResult.rows,
      summaryByUser: summaryByUserResult.rows,
      users,
      invoices: invoicesResult.rows,
      invoiceMonth,
      message,
      dayjs
    });
  } catch (err) {
    console.error(err);
    return res.status(500).send('Erro ao carregar dashboard.');
  }
});

app.post('/withdraw', requireAuth, async (req, res) => {
  const user = req.session.user;

  if (user.role !== 'employee') {
    return res.status(403).send('Somente colaboradores podem registrar retiradas.');
  }

  const { item_id } = req.body;

  try {
    const productResult = await pool.query(
      `SELECT * FROM products WHERE id = $1`,
      [item_id]
    );

    const product = productResult.rows[0];

    if (!product) {
      req.session.message = 'Produto não encontrado.';
      return res.redirect('/dashboard');
    }

    if (Number(product.stock_quantity) <= 0) {
      req.session.message = 'Este produto está sem estoque.';
      return res.redirect('/dashboard');
    }

    await pool.query('BEGIN');

    await pool.query(
      `
      INSERT INTO withdrawals (user_id, item_id, item_name, item_price)
      VALUES ($1, $2, $3, $4)
      `,
      [user.id, product.id, product.name, product.price]
    );

    await pool.query(
      `
      UPDATE products
      SET stock_quantity = stock_quantity - 1
      WHERE id = $1
      `,
      [product.id]
    );

    await pool.query('COMMIT');

    req.session.message = `Retirada registrada com sucesso: ${product.name} - R$ ${Number(product.price).toFixed(2).replace('.', ',')}.`;
    res.redirect('/dashboard');
  } catch (err) {
    await pool.query('ROLLBACK').catch(() => {});
    console.error(err);
    req.session.message = 'Erro ao registrar retirada.';
    res.redirect('/dashboard');
  }
});

app.post('/admin/users', requireAdmin, async (req, res) => {
  const { name, cpf, username, password, role } = req.body;

  if (!name || !cpf || !username || !password || !role) {
    req.session.message = 'Preencha todos os campos do usuário, incluindo o CPF.';
    return res.redirect('/dashboard');
  }

  if (!['admin', 'finance', 'employee'].includes(role)) {
    req.session.message = 'Perfil inválido.';
    return res.redirect('/dashboard');
  }

  // Limpar CPF (só dígitos)
  const cpfLimpo = cpf.replace(/\D/g, '');

  if (cpfLimpo.length !== 11) {
    req.session.message = 'CPF inválido. Informe os 11 dígitos.';
    return res.redirect('/dashboard');
  }

  // Validar CPF (dígitos verificadores)
  function validarCPF(c) {
    if (/^(\d)\1{10}$/.test(c)) return false;
    let soma = 0;
    for (let i = 0; i < 9; i++) soma += Number(c[i]) * (10 - i);
    let resto = (soma * 10) % 11;
    if (resto === 10) resto = 0;
    if (resto !== Number(c[9])) return false;
    soma = 0;
    for (let i = 0; i < 10; i++) soma += Number(c[i]) * (11 - i);
    resto = (soma * 10) % 11;
    if (resto === 10) resto = 0;
    return resto === Number(c[10]);
  }

  if (!validarCPF(cpfLimpo)) {
    req.session.message = 'CPF inválido. Verifique os dígitos informados.';
    return res.redirect('/dashboard');
  }

  // Formatar CPF para salvar padronizado
  const cpfFormatado = cpfLimpo.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');

  try {
    // Verificar se CPF já existe
    const cpfExiste = await pool.query('SELECT id FROM users WHERE cpf = $1', [cpfFormatado]);
    if (cpfExiste.rows.length > 0) {
      req.session.message = 'Já existe um usuário cadastrado com este CPF.';
      return res.redirect('/dashboard');
    }

    const passwordHash = bcrypt.hashSync(password, 10);

    await pool.query(
      `INSERT INTO users (name, cpf, username, password_hash, role) VALUES ($1, $2, $3, $4, $5)`,
      [name.trim(), cpfFormatado, username.trim(), passwordHash, role]
    );

    req.session.message = 'Usuário cadastrado com sucesso.';
    res.redirect('/dashboard');
  } catch (err) {
    console.error(err);
    if (err.code === '23505' && err.constraint && err.constraint.includes('cpf')) {
      req.session.message = 'Já existe um usuário cadastrado com este CPF.';
      return res.redirect('/dashboard');
    }
    req.session.message = 'Não foi possível cadastrar o usuário.';
    res.redirect('/dashboard');
  }
});

app.post('/admin/users/:id/delete', requireAdmin, async (req, res) => {
  const { id } = req.params;

  try {
    // Verificar se o usu\u00e1rio existe e n\u00e3o \u00e9 admin/financeiro fixo
    const userResult = await pool.query('SELECT username, name FROM users WHERE id = $1', [id]);
    const targetUser = userResult.rows[0];

    if (!targetUser) {
      req.session.message = 'Usu\u00e1rio n\u00e3o encontrado.';
      return res.redirect('/dashboard');
    }

    if (targetUser.username === 'admin' || targetUser.username === 'financeiro') {
      req.session.message = 'N\u00e3o \u00e9 poss\u00edvel excluir usu\u00e1rios fixos do sistema.';
      return res.redirect('/dashboard');
    }

    // Excluir usu\u00e1rio (withdrawals ser\u00e3o removidos por CASCADE)
    await pool.query('DELETE FROM users WHERE id = $1', [id]);

    req.session.message = `Usu\u00e1rio "${targetUser.name}" exclu\u00eddo com sucesso.`;
    res.redirect('/dashboard');
  } catch (err) {
    console.error(err);
    req.session.message = 'Erro ao excluir usu\u00e1rio.';
    res.redirect('/dashboard');
  }
});

app.post('/admin/products', requireAdmin, async (req, res) => {
  const { name, price, image_url, stock_quantity } = req.body;

  if (!name || price === undefined || price === '') {
    req.session.message = 'Preencha nome e preço do produto.';
    return res.redirect('/dashboard');
  }

  try {
    await pool.query(
      `INSERT INTO products (name, price, image_url, stock_quantity) VALUES ($1, $2, $3, $4)`,
      [
        name.trim(),
        Number(price),
        image_url ? image_url.trim() : '',
        Number(stock_quantity || 0),
      ]
    );

    req.session.message = 'Produto cadastrado com sucesso.';
    res.redirect('/dashboard');
  } catch (err) {
    console.error(err);
    req.session.message = 'Não foi possível cadastrar o produto.';
    res.redirect('/dashboard');
  }
});

app.post('/admin/products/:id/update', requireAdmin, async (req, res) => {
  const { id } = req.params;
  const { name, price, image_url, stock_quantity } = req.body;

  if (!name || price === undefined || price === '') {
    req.session.message = 'Preencha nome e preço para atualizar.';
    return res.redirect('/dashboard');
  }

  try {
    await pool.query(
      `
      UPDATE products
      SET name = $1, price = $2, image_url = $3, stock_quantity = $4
      WHERE id = $5
      `,
      [
        name.trim(),
        Number(price),
        image_url ? image_url.trim() : '',
        Number(stock_quantity || 0),
        id,
      ]
    );

    req.session.message = 'Produto atualizado com sucesso.';
    res.redirect('/dashboard');
  } catch (err) {
    console.error(err);
    req.session.message = 'Erro ao atualizar produto.';
    res.redirect('/dashboard');
  }
});

app.post('/admin/products/:id/delete', requireAdmin, async (req, res) => {
  const { id } = req.params;

  try {
    const countResult = await pool.query(
      `SELECT COUNT(*) AS total FROM withdrawals WHERE item_id = $1`,
      [id]
    );

    const total = Number(countResult.rows[0]?.total || 0);

    if (total > 0) {
      req.session.message = 'Não é possível excluir um produto que já possui lançamentos.';
      return res.redirect('/dashboard');
    }

    await pool.query(`DELETE FROM products WHERE id = $1`, [id]);

    req.session.message = 'Produto excluído com sucesso.';
    res.redirect('/dashboard');
  } catch (err) {
    console.error(err);
    req.session.message = 'Erro ao excluir produto.';
    res.redirect('/dashboard');
  }
});

app.post('/admin/stock/add', requireFinanceOrAdmin, async (req, res) => {
  const { product_id, quantity, total_cost } = req.body;

  if (!product_id || !quantity || Number(quantity) <= 0) {
    req.session.message = 'Informe um produto e uma quantidade v\u00e1lida.';
    return res.redirect('/dashboard');
  }

  const qty = Number(quantity);
  const totalCost = Number(total_cost || 0);
  const cost = qty > 0 ? totalCost / qty : 0;

  try {
    // Buscar nome do produto
    const prodResult = await pool.query('SELECT name FROM products WHERE id = $1', [product_id]);
    const prodName = prodResult.rows[0] ? prodResult.rows[0].name : 'Desconhecido';

    await pool.query('BEGIN');

    // Atualizar estoque do produto
    await pool.query(
      `UPDATE products SET stock_quantity = stock_quantity + $1 WHERE id = $2`,
      [qty, product_id]
    );

    // Registrar entrada no hist\u00f3rico
    await pool.query(
      `INSERT INTO stock_entries (product_id, product_name, quantity, unit_cost, total_cost)
       VALUES ($1, $2, $3, $4, $5)`,
      [product_id, prodName, qty, cost, totalCost]
    );

    await pool.query('COMMIT');

    req.session.message = `Estoque atualizado: +${qty} ${prodName} (Custo: R$ ${totalCost.toFixed(2).replace('.', ',')})`;
    res.redirect('/dashboard');
  } catch (err) {
    await pool.query('ROLLBACK').catch(() => {});
    console.error(err);
    req.session.message = 'Erro ao atualizar estoque.';
    res.redirect('/dashboard');
  }
});

app.post('/admin/withdrawals/:id/delete', requireAdmin, async (req, res) => {
  const { id } = req.params;

  try {
    await pool.query(`DELETE FROM withdrawals WHERE id = $1`, [id]);

    req.session.message = 'Lançamento excluído com sucesso.';
    res.redirect('/dashboard');
  } catch (err) {
    console.error(err);
    req.session.message = 'Erro ao excluir lançamento.';
    res.redirect('/dashboard');
  }
});

app.get('/reports/xlsx', requireFinanceOrAdmin, async (req, res) => {
  const { start, end } = req.query;

  if (!start || !end || !/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
    return res.status(400).send('Informe a data de in\u00edcio e fim no formato YYYY-MM-DD.');
  }

  if (dayjs(end).isBefore(dayjs(start))) {
    return res.status(400).send('A data fim n\u00e3o pode ser anterior \u00e0 data de in\u00edcio.');
  }

  const endPlusOne = dayjs(end).add(1, 'day').format('YYYY-MM-DD');
  const periodoLabel = `${dayjs(start).format('DD-MM-YYYY')}_a_${dayjs(end).format('DD-MM-YYYY')}`;
  const periodoTitulo = `${dayjs(start).format('DD/MM/YYYY')} a ${dayjs(end).format('DD/MM/YYYY')}`;

  try {
    // Buscar todos os pedidos do per\u00edodo
    const result = await pool.query(
      `SELECT u.name AS "Colaborador", u.username AS "Usuario", w.item_name AS "Item", w.item_price AS "Valor"
       FROM withdrawals w INNER JOIN users u ON u.id = w.user_id
       WHERE w.created_at >= $1 AND w.created_at < $2
       ORDER BY u.name ASC`,
      [start, endPlusOne]
    );
    const rows = result.rows;

    // Buscar entradas de estoque do per\u00edodo
    const stockResult = await pool.query(
      `SELECT product_name, SUM(quantity) AS total_entrada, SUM(total_cost) AS custo_total
       FROM stock_entries
       WHERE created_at >= $1 AND created_at < $2
       GROUP BY product_name ORDER BY product_name ASC`,
      [start, endPlusOne]
    );
    const stockRows = stockResult.rows;

    // Buscar pre\u00e7os de venda dos produtos
    const allProducts = await pool.query('SELECT name, price FROM products ORDER BY name ASC');
    const precosVenda = {};
    for (const p of allProducts.rows) {
      precosVenda[p.name] = Number(p.price);
    }

    // Descobrir todos os produtos distintos e seus pre\u00e7os unit\u00e1rios
    const produtosMap = {};
    for (const row of rows) {
      if (!produtosMap[row.Item]) produtosMap[row.Item] = Number(row.Valor || 0);
    }
    const produtos = Object.keys(produtosMap).sort();
    const precos = produtos.map(p => produtosMap[p]);
    const numProdutos = produtos.length;

    // Agrupar por pessoa
    const pessoas = {};
    for (const row of rows) {
      if (!pessoas[row.Colaborador]) pessoas[row.Colaborador] = { usuario: row.Usuario, produtos: {} };
      if (!pessoas[row.Colaborador].produtos[row.Item]) pessoas[row.Colaborador].produtos[row.Item] = 0;
      pessoas[row.Colaborador].produtos[row.Item] += 1;
    }
    const pessoasOrdenadas = Object.keys(pessoas).sort();

    // Vendas por produto (para aba estoque)
    const vendasPorProduto = {};
    for (const row of rows) {
      if (!vendasPorProduto[row.Item]) vendasPorProduto[row.Item] = { qtd: 0, receita: 0 };
      vendasPorProduto[row.Item].qtd += 1;
      vendasPorProduto[row.Item].receita += Number(row.Valor || 0);
    }

    // ===== GERAR EXCEL COM EXCELJS =====
    const workbook = new ExcelJS.Workbook();

    // Estilos reutiliz\u00e1veis
    const azulClaro = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB8D4E8' } };
    const branco = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    const verdeCl = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
    const bordaFina = {
      top: { style: 'thin', color: { argb: 'FF888888' } },
      left: { style: 'thin', color: { argb: 'FF888888' } },
      bottom: { style: 'thin', color: { argb: 'FF888888' } },
      right: { style: 'thin', color: { argb: 'FF888888' } },
    };
    const fonteNegrito = { bold: true, size: 10, name: 'Arial' };
    const fonteNormal = { size: 10, name: 'Arial' };
    const fonteTitulo = { bold: true, size: 13, name: 'Arial' };
    const fontePreco = { bold: true, size: 9, name: 'Arial', color: { argb: 'FFCC0000' } };
    const fonteTotalVerm = { bold: true, size: 10, name: 'Arial', color: { argb: 'FFCC0000' } };
    const alinhaCentro = { horizontal: 'center', vertical: 'middle', wrapText: true };
    const alinhaEsquerda = { horizontal: 'left', vertical: 'middle', wrapText: true };

    // ============================================================
    // ABA 1: CONTROLE DE CONSUMO
    // ============================================================
    const ws = workbook.addWorksheet('Controle de Consumo');

    // Colunas: Qtde | Nome | C Custo | [qtd produtos...] | [total produtos...] | Total Produtos | Total R$
    const colInicioQtd = 4;
    const colInicioTotal = colInicioQtd + numProdutos;
    const colTotalProdutos = colInicioTotal + numProdutos;
    const colTotalRS = colTotalProdutos + 1;
    const totalCols = colTotalRS;

    // LINHA 1: T\u00cdTULO
    ws.mergeCells(1, 1, 1, totalCols);
    const cellTitulo = ws.getCell(1, 1);
    cellTitulo.value = `CONTROLE DE CONSUMO INTERNO PARA DESCONTO EM FOLHA - ${periodoTitulo}`;
    cellTitulo.font = fonteTitulo;
    cellTitulo.alignment = { horizontal: 'center', vertical: 'middle' };
    cellTitulo.fill = azulClaro;
    ws.getRow(1).height = 30;

    // LINHAS 2-4: CABE\u00c7ALHO
    ws.mergeCells(2, 1, 4, 1);
    ws.mergeCells(2, 2, 4, 2);
    ws.mergeCells(2, 3, 4, 3);

    const setCellStyle = (r, c, val, font, align, fill) => {
      const cell = ws.getCell(r, c);
      cell.value = val; cell.font = font; cell.alignment = align; cell.fill = fill; cell.border = bordaFina;
      return cell;
    };

    setCellStyle(2, 1, 'Qtde', fonteNegrito, alinhaCentro, azulClaro);
    setCellStyle(2, 2, 'NOME', fonteNegrito, alinhaCentro, azulClaro);
    setCellStyle(2, 3, 'C CUSTO', fonteNegrito, alinhaCentro, azulClaro);

    for (let i = 0; i < numProdutos; i++) {
      const col = colInicioQtd + i;
      ws.mergeCells(2, col, 3, col);
      setCellStyle(2, col, produtos[i].toUpperCase(), fonteNegrito, alinhaCentro, azulClaro);
      setCellStyle(4, col, `R$  ${precos[i].toFixed(2).replace('.', ',')}`, fontePreco, alinhaCentro, azulClaro);
    }

    // Cabe\u00e7alho TOTAL (produtos R$)
    if (numProdutos > 0) {
      ws.mergeCells(2, colInicioTotal, 2, colInicioTotal + numProdutos - 1);
      setCellStyle(2, colInicioTotal, 'TOTAL', { bold: true, size: 12, name: 'Arial', color: { argb: 'FFCC0000' } }, alinhaCentro, azulClaro);
    }
    for (let i = 0; i < numProdutos; i++) {
      const col = colInicioTotal + i;
      ws.mergeCells(3, col, 4, col);
      setCellStyle(3, col, produtos[i].toUpperCase(), fonteNegrito, alinhaCentro, azulClaro);
    }

    // Cabe\u00e7alhos Total Produtos e Total R$
    ws.mergeCells(2, colTotalProdutos, 4, colTotalProdutos);
    setCellStyle(2, colTotalProdutos, 'TOTAL PRODUTOS', fonteTotalVerm, alinhaCentro, azulClaro);
    ws.mergeCells(2, colTotalRS, 4, colTotalRS);
    setCellStyle(2, colTotalRS, 'TOTAL R$', fonteTotalVerm, alinhaCentro, azulClaro);

    // Bordas nas c\u00e9lulas merged
    for (let r = 3; r <= 4; r++) {
      for (let c = 1; c <= 3; c++) {
        ws.getCell(r, c).border = bordaFina;
        ws.getCell(r, c).fill = azulClaro;
      }
    }

    // LINHAS DE DADOS
    let linhaAtual = 5;
    for (let p = 0; p < pessoasOrdenadas.length; p++) {
      const nome = pessoasOrdenadas[p];
      const dados = pessoas[nome];
      const fillLinha = p % 2 === 0 ? azulClaro : branco;

      setCellStyle(linhaAtual, 1, 1, fonteNormal, alinhaCentro, fillLinha);
      setCellStyle(linhaAtual, 2, nome.toUpperCase(), fonteNormal, alinhaEsquerda, fillLinha);
      setCellStyle(linhaAtual, 3, dados.usuario ? dados.usuario.toUpperCase() : '-', fonteNormal, alinhaEsquerda, fillLinha);

      let totalProdPessoa = 0;
      let totalRSPessoa = 0;

      for (let i = 0; i < numProdutos; i++) {
        const col = colInicioQtd + i;
        const qtd = dados.produtos[produtos[i]] || 0;
        totalProdPessoa += qtd;
        setCellStyle(linhaAtual, col, qtd > 0 ? qtd : '-', fonteNormal, alinhaCentro, fillLinha);
      }

      for (let i = 0; i < numProdutos; i++) {
        const col = colInicioTotal + i;
        const qtd = dados.produtos[produtos[i]] || 0;
        const totalProduto = qtd * precos[i];
        totalRSPessoa += totalProduto;
        const cell = setCellStyle(linhaAtual, col, totalProduto > 0 ? Number(totalProduto.toFixed(2)) : '-', fonteNormal, alinhaCentro, fillLinha);
        if (totalProduto > 0) cell.numFmt = '#,##0.00';
      }

      // Total Produtos
      const cTP = setCellStyle(linhaAtual, colTotalProdutos, totalProdPessoa, fonteTotalVerm, alinhaCentro, fillLinha);
      // Total R$
      const cTR = setCellStyle(linhaAtual, colTotalRS, Number(totalRSPessoa.toFixed(2)), fonteTotalVerm, alinhaCentro, fillLinha);
      cTR.numFmt = '#,##0.00';

      linhaAtual++;
    }

    // LARGURAS
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 42;
    ws.getColumn(3).width = 20;
    for (let i = 0; i < numProdutos; i++) {
      ws.getColumn(colInicioQtd + i).width = 16;
      ws.getColumn(colInicioTotal + i).width = 16;
    }
    ws.getColumn(colTotalProdutos).width = 16;
    ws.getColumn(colTotalRS).width = 16;
    ws.getRow(2).height = 28;
    ws.getRow(3).height = 22;
    ws.getRow(4).height = 20;

    // ============================================================
    // ABA 2: ESTOQUE
    // ============================================================
    const wsE = workbook.addWorksheet('Estoque');

    const colsEstoque = 7;
    // Colunas: Produto | Pre\u00e7o Venda | Qtd Entrada | Qtd Vendida | Custo Total | Receita Venda | Lucro

    // LINHA 1: T\u00cdTULO
    wsE.mergeCells(1, 1, 1, colsEstoque);
    const cellTituloE = wsE.getCell(1, 1);
    cellTituloE.value = `CONTROLE DE ESTOQUE - ${periodoTitulo}`;
    cellTituloE.font = fonteTitulo;
    cellTituloE.alignment = { horizontal: 'center', vertical: 'middle' };
    cellTituloE.fill = azulClaro;
    wsE.getRow(1).height = 30;

    // LINHA 2: CABE\u00c7ALHO
    const cabEstoque = ['PRODUTO', 'PRE\u00c7O VENDA (R$)', 'QTD ENTRADA', 'QTD VENDIDA', 'CUSTO TOTAL (R$)', 'RECEITA VENDA (R$)', 'LUCRO (R$)'];
    for (let c = 0; c < cabEstoque.length; c++) {
      const cell = wsE.getCell(2, c + 1);
      cell.value = cabEstoque[c];
      cell.font = fonteNegrito;
      cell.alignment = alinhaCentro;
      cell.fill = azulClaro;
      cell.border = bordaFina;
    }
    wsE.getRow(2).height = 28;

    // Juntar dados de estoque e vendas por produto
    const todosProdutosEstoque = new Set();
    for (const s of stockRows) todosProdutosEstoque.add(s.product_name);
    for (const v of Object.keys(vendasPorProduto)) todosProdutosEstoque.add(v);
    const produtosEstoqueOrdenados = [...todosProdutosEstoque].sort();

    let linhaE = 3;
    let totalEntradaGeral = 0, totalVendidaGeral = 0, totalCustoGeral = 0, totalReceitaGeral = 0, totalLucroGeral = 0;

    for (let p = 0; p < produtosEstoqueOrdenados.length; p++) {
      const nome = produtosEstoqueOrdenados[p];
      const fillLinha = p % 2 === 0 ? azulClaro : branco;

      const stockInfo = stockRows.find(s => s.product_name === nome);
      const vendaInfo = vendasPorProduto[nome];
      const precoVenda = precosVenda[nome] || 0;

      const qtdEntrada = stockInfo ? Number(stockInfo.total_entrada) : 0;
      const custoTotal = stockInfo ? Number(stockInfo.custo_total) : 0;
      const qtdVendida = vendaInfo ? vendaInfo.qtd : 0;
      const receitaVenda = vendaInfo ? vendaInfo.receita : 0;
      const lucro = receitaVenda - custoTotal;

      totalEntradaGeral += qtdEntrada;
      totalVendidaGeral += qtdVendida;
      totalCustoGeral += custoTotal;
      totalReceitaGeral += receitaVenda;
      totalLucroGeral += lucro;

      const setE = (c, val, fmt) => {
        const cell = wsE.getCell(linhaE, c);
        cell.value = val; cell.font = fonteNormal; cell.alignment = alinhaCentro; cell.fill = fillLinha; cell.border = bordaFina;
        if (fmt) cell.numFmt = fmt;
        return cell;
      };

      wsE.getCell(linhaE, 1).value = nome.toUpperCase();
      wsE.getCell(linhaE, 1).font = fonteNormal;
      wsE.getCell(linhaE, 1).alignment = alinhaEsquerda;
      wsE.getCell(linhaE, 1).fill = fillLinha;
      wsE.getCell(linhaE, 1).border = bordaFina;

      setE(2, Number(precoVenda.toFixed(2)), '#,##0.00');
      setE(3, qtdEntrada);
      setE(4, qtdVendida);
      setE(5, Number(custoTotal.toFixed(2)), '#,##0.00');
      setE(6, Number(receitaVenda.toFixed(2)), '#,##0.00');

      const cellLucro = setE(7, Number(lucro.toFixed(2)), '#,##0.00');
      cellLucro.font = { bold: true, size: 10, name: 'Arial', color: { argb: lucro >= 0 ? 'FF006100' : 'FFCC0000' } };

      linhaE++;
    }

    // LINHA TOTAL GERAL
    const fillTotal = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9CC0DE' } };
    const setTot = (c, val, fmt) => {
      const cell = wsE.getCell(linhaE, c);
      cell.value = val; cell.font = fonteTotalVerm; cell.alignment = alinhaCentro; cell.fill = fillTotal; cell.border = bordaFina;
      if (fmt) cell.numFmt = fmt;
      return cell;
    };
    wsE.getCell(linhaE, 1).value = 'TOTAL GERAL';
    wsE.getCell(linhaE, 1).font = fonteTotalVerm;
    wsE.getCell(linhaE, 1).alignment = alinhaEsquerda;
    wsE.getCell(linhaE, 1).fill = fillTotal;
    wsE.getCell(linhaE, 1).border = bordaFina;
    setTot(2, '-');
    setTot(3, totalEntradaGeral);
    setTot(4, totalVendidaGeral);
    setTot(5, Number(totalCustoGeral.toFixed(2)), '#,##0.00');
    setTot(6, Number(totalReceitaGeral.toFixed(2)), '#,##0.00');
    const cellLucroTotal = setTot(7, Number(totalLucroGeral.toFixed(2)), '#,##0.00');
    cellLucroTotal.font = { bold: true, size: 11, name: 'Arial', color: { argb: totalLucroGeral >= 0 ? 'FF006100' : 'FFCC0000' } };

    // LARGURAS ABA ESTOQUE
    wsE.getColumn(1).width = 35;
    wsE.getColumn(2).width = 18;
    wsE.getColumn(3).width = 16;
    wsE.getColumn(4).width = 16;
    wsE.getColumn(5).width = 20;
    wsE.getColumn(6).width = 20;
    wsE.getColumn(7).width = 18;

    // ===== SALVAR E ENVIAR =====
    const fileName = `controle-consumo-${periodoLabel}.xlsx`;
    const filePath = path.join(baseDir, fileName);
    await workbook.xlsx.writeFile(filePath);

    res.download(filePath, fileName, (downloadErr) => {
      if (downloadErr) console.error(downloadErr);
      if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    });
  } catch (err) {
    console.error(err);
    res.status(500).send('Erro ao gerar relat\u00f3rio.');
  }
});

// ===== ROTAS DE NOTAS FISCAIS =====

app.post('/invoices/upload', requireFinanceOrAdmin, upload.single('invoice_file'), async (req, res) => {
  const { reference_month } = req.body;
  const file = req.file;

  if (!reference_month || !file) {
    req.session.message = 'Informe o m\u00eas de refer\u00eancia e selecione um arquivo.';
    return res.redirect('/dashboard');
  }

  try {
    await pool.query(
      `INSERT INTO invoices (reference_month, file_name, file_type, file_data, uploaded_by)
       VALUES ($1, $2, $3, $4, $5)`,
      [reference_month, file.originalname, file.mimetype, file.buffer, req.session.user.id]
    );

    req.session.message = `Nota fiscal "${file.originalname}" enviada com sucesso para ${reference_month}.`;
    res.redirect('/dashboard');
  } catch (err) {
    console.error(err);
    req.session.message = 'Erro ao enviar nota fiscal.';
    res.redirect('/dashboard');
  }
});

app.get('/invoices/:id/view', requireFinanceOrAdmin, async (req, res) => {
  const { id } = req.params;

  try {
    const result = await pool.query('SELECT file_name, file_type, file_data FROM invoices WHERE id = $1', [id]);
    const invoice = result.rows[0];

    if (!invoice) return res.status(404).send('Nota fiscal n\u00e3o encontrada.');

    res.setHeader('Content-Type', invoice.file_type);
    res.setHeader('Content-Disposition', `inline; filename="${invoice.file_name}"`);
    res.send(invoice.file_data);
  } catch (err) {
    console.error(err);
    res.status(500).send('Erro ao visualizar nota fiscal.');
  }
});

app.get('/invoices/:id/download', requireFinanceOrAdmin, async (req, res) => {
  const { id } = req.params;

  try {
    const result = await pool.query('SELECT file_name, file_type, file_data FROM invoices WHERE id = $1', [id]);
    const invoice = result.rows[0];

    if (!invoice) return res.status(404).send('Nota fiscal n\u00e3o encontrada.');

    res.setHeader('Content-Type', invoice.file_type);
    res.setHeader('Content-Disposition', `attachment; filename="${invoice.file_name}"`);
    res.send(invoice.file_data);
  } catch (err) {
    console.error(err);
    res.status(500).send('Erro ao baixar nota fiscal.');
  }
});

app.post('/invoices/:id/delete', requireFinanceOrAdmin, async (req, res) => {
  const { id } = req.params;

  try {
    const result = await pool.query('SELECT file_name FROM invoices WHERE id = $1', [id]);
    const invoice = result.rows[0];

    if (!invoice) {
      req.session.message = 'Nota fiscal n\u00e3o encontrada.';
      return res.redirect('/dashboard');
    }

    await pool.query('DELETE FROM invoices WHERE id = $1', [id]);

    req.session.message = `Nota fiscal "${invoice.file_name}" exclu\u00edda com sucesso.`;
    res.redirect('/dashboard');
  } catch (err) {
    console.error(err);
    req.session.message = 'Erro ao excluir nota fiscal.';
    res.redirect('/dashboard');
  }
});

app.get('/logout', (req, res) => {
  req.session.destroy(() => {
    res.redirect('/');
  });
});

initDB()
  .then(() => {
    app.listen(PORT, '0.0.0.0', () => {
      console.log(`Servidor rodando em http://0.0.0.0:${PORT}`);
    });
  })
  .catch((err) => {
    console.error('Erro ao inicializar banco:', err);
  });
