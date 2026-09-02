/* =========================================================
   SISTEMA DE VENTAS E INSTALACIONES - SUPABASE
   Frontend estático para GitHub Pages.
   La seguridad real se aplica mediante RLS en Supabase.
   ========================================================= */

(function () {
"use strict";

if (window.__ventasInstalacionesAppLoaded) return;
window.__ventasInstalacionesAppLoaded = true;

const SUPABASE_URL = "https://jsyeczuhdjusbcmpiiyg.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_5gFuPfsCqONtLc1G_gk-jQ_eUPK30zp";

const { createClient } = window.supabase;
const sbClient = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY);

let currentUser = null;
let currentProfile = null;
let sales = [];
let advisors = [];

document.addEventListener("DOMContentLoaded", async () => {
  bindEvents();
  setTodayDefault();
  showAuthView();

  const { data: { session } } = await sbClient.auth.getSession();

  if (session?.user) {
    await initializeSession(session.user);
  }

  sbClient.auth.onAuthStateChange(async (event, session) => {
    if (event === "SIGNED_OUT") {
      currentUser = null;
      currentProfile = null;
      sales = [];
      advisors = [];
      showAuthView();
      return;
    }

    if (session?.user && event !== "INITIAL_SESSION") {
      await initializeSession(session.user);
    }
  });
});

function bindEvents() {
  document.getElementById("login-form").addEventListener("submit", login);
  document.getElementById("register-form").addEventListener("submit", registerAdvisor);
  document.getElementById("sale-form").addEventListener("submit", registerSale);

  document.getElementById("btn-show-register").addEventListener("click", () => {
    document.getElementById("auth-view").classList.add("hidden");
    document.getElementById("register-view").classList.remove("hidden");
  });

  document.getElementById("btn-back-login").addEventListener("click", showAuthView);
  document.getElementById("btn-logout").addEventListener("click", logout);

  document.getElementById("filtroAsesor").addEventListener("input", renderAdvisorTable);

  [
    "filtroAdminTexto",
    "filtroAsesorAdmin",
    "filtroEstadoAdmin",
    "filtroTipoAdmin",
    "filtroZonaAdmin",
    "filtroDesdeAdmin",
    "filtroHastaAdmin"
  ].forEach((id) => {
    document.getElementById(id).addEventListener("input", renderAdmin);
    document.getElementById(id).addEventListener("change", renderAdmin);
  });

  document.getElementById("btn-clear-filters").addEventListener("click", clearAdminFilters);
  document.getElementById("btn-print-report").addEventListener("click", printReport);
}

async function login(event) {
  event.preventDefault();

  const email = value("login-email");
  const password = value("login-password");

  if (!email || !password) return;

  setButtonBusy(event.submitter, true, "Ingresando...");

  const { data, error } = await sbClient.auth.signInWithPassword({
    email,
    password
  });

  setButtonBusy(event.submitter, false, "Ingresar");

  if (error) {
    showToast(authError(error), true);
    return;
  }

  await initializeSession(data.user);
}

async function registerAdvisor(event) {
  event.preventDefault();

  const nombre = value("reg-nombre");
  const apellido = value("reg-apellido");
  const documento = value("reg-documento");
  const telefono = value("reg-telefono");
  const zona = value("reg-zona");
  const email = value("reg-email");
  const password = document.getElementById("reg-password").value;
  const confirm = document.getElementById("reg-password-confirm").value;

  if (password !== confirm) {
    showToast("Las contraseñas no coinciden.", true);
    return;
  }

  if (password.length < 6) {
    showToast("La contraseña debe tener mínimo 6 caracteres.", true);
    return;
  }

  setButtonBusy(event.submitter, true, "Registrando...");

  const { data, error } = await sbClient.auth.signUp({
    email,
    password,
    options: {
      data: {
        nombre,
        apellido,
        documento,
        telefono,
        zona,
        rol: "asesor"
      }
    }
  });

  setButtonBusy(event.submitter, false, "Registrar asesor");

  if (error) {
    showToast(authError(error), true);
    return;
  }

  document.getElementById("register-form").reset();

  if (data.session) {
    showToast("Asesor registrado correctamente.");
    await initializeSession(data.user);
  } else {
    showToast("Registro creado. Revisa tu correo para confirmar la cuenta.");
    showAuthView();
  }
}

async function initializeSession(user) {
  currentUser = user;

  const { data: profile, error } = await sbClient
    .from("perfiles")
    .select("*")
    .eq("id", user.id)
    .single();

  if (error) {
    console.error(error);
    await sbClient.auth.signOut();
    showToast("No fue posible cargar tu perfil. Verifica que hayas ejecutado el SQL de configuración.", true);
    return;
  }

  currentProfile = profile;
  updateSessionHeader();

  if (profile.rol === "administrador") {
    await loadAdminData();
    showView("vista-admin");
  } else {
    await loadAdvisorData();
    showView("vista-asesor");
  }
}

async function loadAdvisorData() {
  const { data, error } = await sbClient
    .from("ventas")
    .select("*")
    .eq("asesor_id", currentUser.id)
    .order("fecha_venta", { ascending: false })
    .order("id", { ascending: false });

  if (error) {
    console.error(error);
    showToast("No fue posible cargar tus operaciones.", true);
    return;
  }

  sales = data || [];
  applyAdvisorProfile();
  renderAdvisorTable();
}

async function loadAdminData() {
  const [salesResult, advisorsResult] = await Promise.all([
    sbClient
      .from("ventas")
      .select(`
        *,
        perfiles:asesor_id (
          id,
          nombre,
          apellido,
          zona,
          email
        )
      `)
      .order("fecha_venta", { ascending: false })
      .order("id", { ascending: false }),

    sbClient
      .from("perfiles")
      .select("id,nombre,apellido,zona,email")
      .eq("rol", "asesor")
      .order("nombre", { ascending: true })
      .order("apellido", { ascending: true })
  ]);

  if (salesResult.error) {
    console.error(salesResult.error);
    showToast("No fue posible cargar las operaciones administrativas.", true);
    return;
  }

  if (advisorsResult.error) {
    console.error(advisorsResult.error);
    showToast("No fue posible cargar los asesores.", true);
    return;
  }

  sales = salesResult.data || [];
  advisors = advisorsResult.data || [];

  populateAdminFilters();
  renderAdmin();
}

async function registerSale(event) {
  event.preventDefault();

  if (!currentUser || !currentProfile) {
    showToast("Tu sesión no está disponible.", true);
    return;
  }

  const tipoOperacion = value("tipoOperacion");
  const codigoCliente = value("codigoCliente");
  const codigoServicio = value("codigoServicio");
  const descripcionServicio = value("descripcionServicio");
  const zona = value("zona");
  const fechaVenta = document.getElementById("fechaVenta").value;

  if (!tipoOperacion || !codigoCliente || !codigoServicio || !descripcionServicio || !zona || !fechaVenta) {
    showToast("Completa todos los campos obligatorios.", true);
    return;
  }

  const { data, error } = await sbClient
    .from("ventas")
    .insert({
      asesor_id: currentUser.id,
      tipo_operacion: tipoOperacion,
      codigo_cliente: codigoCliente,
      codigo_servicio: codigoServicio,
      descripcion_servicio: descripcionServicio,
      zona,
      fecha_venta: fechaVenta,
      estado_instalacion: "PENDIENTE",
      fecha_instalacion: null
    })
    .select()
    .single();

  if (error) {
    console.error(error);
    showToast(error.message || "No fue posible registrar la operación.", true);
    return;
  }

  event.target.reset();
  applyAdvisorProfile();
  setTodayDefault();
  sales.unshift(data);
  renderAdvisorTable();

  showToast(`${tipoOperacion} registrada correctamente. Estado: Instalación pendiente.`);
}

async function setInstallation(id) {
  if (!currentProfile || currentProfile.rol !== "administrador") {
    showToast("No tienes permisos para registrar instalaciones.", true);
    return;
  }

  const input = document.getElementById(`date-${id}`);
  if (!input || !input.value) {
    showToast("Selecciona la fecha en que se realizó la instalación.", true);
    return;
  }

  const { data, error } = await sbClient
    .from("ventas")
    .update({
      fecha_instalacion: input.value,
      estado_instalacion: "REALIZADA"
    })
    .eq("id", id)
    .select(`
      *,
      perfiles:asesor_id (
        id,
        nombre,
        apellido,
        zona,
        email
      )
    `)
    .single();

  if (error) {
    console.error(error);
    showToast("No fue posible actualizar la instalación.", true);
    return;
  }

  const index = sales.findIndex((item) => item.id === id);
  if (index >= 0) sales[index] = data;

  renderAdmin();
  showToast("Instalación marcada como realizada.");
}

async function logout() {
  const { error } = await sbClient.auth.signOut();

  if (error) {
    console.error(error);
    showToast("No fue posible cerrar la sesión.", true);
  }
}

function updateSessionHeader() {
  const sessionArea = document.getElementById("session-area");
  const name = [currentProfile?.nombre, currentProfile?.apellido].filter(Boolean).join(" ") || "Usuario";
  const role = currentProfile?.rol === "administrador" ? "Administrador" : "Asesor";

  document.getElementById("user-name").textContent = name;
  document.getElementById("user-role").textContent = role;
  document.getElementById("user-avatar").textContent = name.charAt(0).toUpperCase();

  sessionArea.classList.remove("hidden");
}

function applyAdvisorProfile() {
  if (!currentProfile) return;

  const zona = currentProfile.zona || "";
  document.getElementById("zona").value = zona;
  document.getElementById("asesor-zone-badge").textContent = `Zona: ${zona || "Sin asignar"}`;
  document.getElementById("asesor-welcome").textContent =
    `Registra ventas y reconexiones. Zona asignada: ${zona || "sin asignar"}.`;
}

function renderAdvisorTable() {
  const tabla = document.getElementById("tabla-asesor");
  const filtro = value("filtroAsesor").toLowerCase();

  const filtered = sales.filter((sale) => {
    const text = [
      sale.tipo_operacion,
      sale.codigo_cliente,
      sale.codigo_servicio,
      sale.descripcion_servicio,
      sale.zona,
      sale.fecha_venta,
      sale.estado_instalacion,
      sale.fecha_instalacion || ""
    ].join(" ").toLowerCase();

    return text.includes(filtro);
  });

  tabla.innerHTML = "";

  if (!filtered.length) {
    tabla.innerHTML = `<tr class="empty-row"><td colspan="9">${
      sales.length ? "No se encontraron operaciones con ese criterio." : "No hay operaciones registradas."
    }</td></tr>`;
  } else {
    filtered.forEach((sale) => {
      const row = document.createElement("tr");
      row.innerHTML = `
        <td>#${sale.id}</td>
        <td>${operationBadge(sale.tipo_operacion)}</td>
        <td>${escapeHTML(sale.codigo_cliente)}</td>
        <td>${escapeHTML(sale.codigo_servicio)}</td>
        <td>${escapeHTML(sale.descripcion_servicio)}</td>
        <td>${escapeHTML(sale.zona)}</td>
        <td>${formatDate(sale.fecha_venta)}</td>
        <td>${installationStatus(sale.estado_instalacion)}</td>
        <td>${formatDate(sale.fecha_instalacion)}</td>
      `;
      tabla.appendChild(row);
    });
  }

  updateAdvisorStats();
}

function renderAdmin() {
  const tabla = document.getElementById("tabla-admin");
  if (!tabla || !currentProfile || currentProfile.rol !== "administrador") return;

  const filtered = getFilteredAdminSales();

  tabla.innerHTML = "";

  if (!filtered.length) {
    tabla.innerHTML = `<tr class="empty-row"><td colspan="11">${
      sales.length ? "No se encontraron operaciones con los filtros seleccionados." : "No hay operaciones registradas."
    }</td></tr>`;
  } else {
    filtered.forEach((sale) => {
      const advisor = sale.perfiles || {};
      const advisorName = [advisor.nombre, advisor.apellido].filter(Boolean).join(" ") || "—";

      const row = document.createElement("tr");
      row.innerHTML = `
        <td>#${sale.id}</td>
        <td>${escapeHTML(advisorName)}</td>
        <td>${operationBadge(sale.tipo_operacion)}</td>
        <td>${escapeHTML(sale.codigo_cliente)}</td>
        <td>${escapeHTML(sale.codigo_servicio)}</td>
        <td>${escapeHTML(sale.descripcion_servicio)}</td>
        <td>${escapeHTML(sale.zona)}</td>
        <td>${formatDate(sale.fecha_venta)}</td>
        <td>${installationStatus(sale.estado_instalacion)}</td>
        <td>
          <input class="installation-date" type="date"
                 id="date-${sale.id}"
                 value="${sale.fecha_instalacion || ""}"
                 aria-label="Fecha de instalación">
        </td>
        <td>
          <button type="button" class="btn-save-installation" onclick="setInstallation(${sale.id})">
            Marcar realizada
          </button>
        </td>
      `;
      tabla.appendChild(row);
    });
  }

  updateAdminStats();
  document.getElementById("admin-result-count").textContent =
    `${filtered.length} resultado${filtered.length === 1 ? "" : "s"}`;
}

function getFilteredAdminSales() {
  const text = value("filtroAdminTexto").toLowerCase();
  const advisorId = document.getElementById("filtroAsesorAdmin").value;
  const estado = document.getElementById("filtroEstadoAdmin").value;
  const tipo = document.getElementById("filtroTipoAdmin").value;
  const zona = document.getElementById("filtroZonaAdmin").value;
  const desde = document.getElementById("filtroDesdeAdmin").value;
  const hasta = document.getElementById("filtroHastaAdmin").value;

  return sales.filter((sale) => {
    const advisor = sale.perfiles || {};
    const searchable = [
      advisor.nombre,
      advisor.apellido,
      advisor.email,
      sale.tipo_operacion,
      sale.codigo_cliente,
      sale.codigo_servicio,
      sale.descripcion_servicio,
      sale.zona,
      sale.fecha_venta
    ].join(" ").toLowerCase();

    if (text && !searchable.includes(text)) return false;
    if (advisorId && sale.asesor_id !== advisorId) return false;
    if (estado && sale.estado_instalacion !== estado) return false;
    if (tipo && sale.tipo_operacion !== tipo) return false;
    if (zona && sale.zona !== zona) return false;
    if (desde && sale.fecha_venta < desde) return false;
    if (hasta && sale.fecha_venta > hasta) return false;

    return true;
  });
}

function populateAdminFilters() {
  const advisorSelect = document.getElementById("filtroAsesorAdmin");
  const zoneSelect = document.getElementById("filtroZonaAdmin");

  const currentAdvisor = advisorSelect.value;
  const currentZone = zoneSelect.value;

  advisorSelect.innerHTML = `<option value="">Todos los asesores</option>`;
  advisors.forEach((advisor) => {
    const option = document.createElement("option");
    option.value = advisor.id;
    option.textContent = [advisor.nombre, advisor.apellido].filter(Boolean).join(" ") || advisor.email || advisor.id;
    advisorSelect.appendChild(option);
  });
  advisorSelect.value = currentAdvisor;

  const zones = [...new Set(sales.map((sale) => sale.zona).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  zoneSelect.innerHTML = `<option value="">Todas las zonas</option>`;
  zones.forEach((zone) => {
    const option = document.createElement("option");
    option.value = zone;
    option.textContent = zone;
    zoneSelect.appendChild(option);
  });
  zoneSelect.value = currentZone;
}

function updateAdvisorStats() {
  const total = sales.length;
  const ventas = sales.filter((s) => s.tipo_operacion === "Venta").length;
  const reconexiones = sales.filter((s) => s.tipo_operacion === "Reconexión").length;
  const completas = sales.filter((s) => s.estado_instalacion === "REALIZADA").length;

  document.getElementById("asesor-total-count").textContent = total;
  document.getElementById("asesor-ventas-count").textContent = ventas;
  document.getElementById("asesor-reconexion-count").textContent = reconexiones;
  document.getElementById("asesor-complete-count").textContent = completas;
}

function updateAdminStats() {
  const total = sales.length;
  const ventas = sales.filter((s) => s.tipo_operacion === "Venta").length;
  const reconexiones = sales.filter((s) => s.tipo_operacion === "Reconexión").length;
  const pendientes = sales.filter((s) => s.estado_instalacion === "PENDIENTE").length;
  const completas = sales.filter((s) => s.estado_instalacion === "REALIZADA").length;

  document.getElementById("admin-total-count").textContent = total;
  document.getElementById("admin-ventas-count").textContent = ventas;
  document.getElementById("admin-reconexion-count").textContent = reconexiones;
  document.getElementById("admin-pending-count").textContent = pendientes;
  document.getElementById("admin-complete-count").textContent = completas;
}

function clearAdminFilters() {
  [
    "filtroAdminTexto",
    "filtroAsesorAdmin",
    "filtroEstadoAdmin",
    "filtroTipoAdmin",
    "filtroZonaAdmin",
    "filtroDesdeAdmin",
    "filtroHastaAdmin"
  ].forEach((id) => {
    const element = document.getElementById(id);
    element.value = "";
  });

  renderAdmin();
}

function printReport() {
  const filtered = getFilteredAdminSales();
  const printArea = document.getElementById("print-report");

  const rows = filtered.map((sale) => {
    const advisor = sale.perfiles || {};
    const advisorName = [advisor.nombre, advisor.apellido].filter(Boolean).join(" ") || "—";

    return `
      <tr>
        <td>#${sale.id}</td>
        <td>${escapeHTML(advisorName)}</td>
        <td>${escapeHTML(sale.tipo_operacion)}</td>
        <td>${escapeHTML(sale.codigo_cliente)}</td>
        <td>${escapeHTML(sale.codigo_servicio)}</td>
        <td>${escapeHTML(sale.zona)}</td>
        <td>${formatDate(sale.fecha_venta)}</td>
        <td>${sale.estado_instalacion === "REALIZADA" ? "Completa" : "Pendiente"}</td>
        <td>${formatDate(sale.fecha_instalacion)}</td>
      </tr>
    `;
  }).join("");

  printArea.innerHTML = `
    <div class="print-header">
      <h1>Reporte de ventas e instalaciones</h1>
      <p>Generado: ${new Date().toLocaleString("es-CO")}</p>
      <p>Total filtrado: <strong>${filtered.length}</strong></p>
    </div>
    <table>
      <thead>
        <tr>
          <th>ID</th>
          <th>Asesor</th>
          <th>Operación</th>
          <th>Cliente</th>
          <th>Servicio</th>
          <th>Zona</th>
          <th>Fecha venta</th>
          <th>Estado</th>
          <th>Fecha instalación</th>
        </tr>
      </thead>
      <tbody>${rows || `<tr><td colspan="9">No hay registros con los filtros seleccionados.</td></tr>`}</tbody>
    </table>
  `;

  window.print();
}

function showAuthView() {
  document.getElementById("auth-view").classList.remove("hidden");
  document.getElementById("register-view").classList.add("hidden");
  document.getElementById("vista-asesor").classList.add("hidden");
  document.getElementById("vista-admin").classList.add("hidden");
  document.getElementById("session-area").classList.add("hidden");
}

function showView(viewId) {
  document.getElementById("auth-view").classList.add("hidden");
  document.getElementById("register-view").classList.add("hidden");
  document.getElementById("vista-asesor").classList.add("hidden");
  document.getElementById("vista-admin").classList.add("hidden");
  document.getElementById(viewId).classList.remove("hidden");
}

function installationStatus(status) {
  return status === "REALIZADA"
    ? `<span class="badge badge-complete">Instalación realizada</span>`
    : `<span class="badge badge-pending">Instalación pendiente</span>`;
}

function operationBadge(tipo) {
  return tipo === "Reconexión"
    ? `<span class="badge badge-reconnection">Reconexión</span>`
    : `<span class="badge badge-sale">Venta</span>`;
}

function formatDate(isoDate) {
  if (!isoDate) return "—";
  const parts = isoDate.split("-");
  return parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : escapeHTML(isoDate);
}

function setTodayDefault() {
  const input = document.getElementById("fechaVenta");
  if (input && !input.value) input.value = getTodayISO();
}

function getTodayISO() {
  const now = new Date();
  const offset = now.getTimezoneOffset();
  const local = new Date(now.getTime() - offset * 60000);
  return local.toISOString().slice(0, 10);
}

function value(id) {
  return document.getElementById(id).value.trim();
}

function escapeHTML(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function setButtonBusy(button, busy, text) {
  if (!button) return;
  button.disabled = busy;
  button.textContent = text;
}

function authError(error) {
  const message = (error?.message || "").toLowerCase();

  if (message.includes("invalid login credentials")) return "Correo o contraseña incorrectos.";
  if (message.includes("email not confirmed")) return "Debes confirmar tu correo antes de iniciar sesión.";
  if (message.includes("user already registered")) return "Ese correo ya está registrado.";
  if (message.includes("password")) return "La contraseña no cumple los requisitos.";
  return error?.message || "No fue posible completar la operación.";
}

let toastTimer;
function showToast(message, isError = false) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.classList.toggle("error", isError);
  toast.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("show"), 3500);
}

})();
