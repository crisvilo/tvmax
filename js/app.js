/* =========================================================
   SISTEMA DE VENTAS E INSTALACIONES - SUPABASE
   ========================================================= */
(function () {
  "use strict";
  if (window.__ventasInstalacionesAppLoaded) return;
  window.__ventasInstalacionesAppLoaded = true;

  const SUPABASE_URL = "https://jsyeczuhdjusbcmpiiyg.supabase.co";
  const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_5gFuPfsCqONtLc1G_gk-jQ_eUPK30zp";
  const { createClient } = window.supabase;
  const sbClient = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY);
  const SERVICES = ["Internet", "TV", "Combo", "Otros"];
  const STATES = ["PENDIENTE", "REALIZADA", "CANCELADA"];
  let currentUser = null, currentProfile = null, sales = [], advisors = [], config = { color_principal: "#8b5cf6", logo_url: "" };

  document.addEventListener("DOMContentLoaded", async () => {
    bindEvents(); setTodayDefault(); showAuthView(); applyTheme();
    const { data: { session } } = await sbClient.auth.getSession();
    if (session?.user) await initializeSession(session.user);
    sbClient.auth.onAuthStateChange(async (event, session) => {
      if (event === "SIGNED_OUT") { currentUser = null; currentProfile = null; sales = []; advisors = []; showAuthView(); return; }
      if (session?.user && event !== "INITIAL_SESSION") await initializeSession(session.user);
    });
  });

  function bindEvents() {
    id("login-form").addEventListener("submit", login); id("register-form").addEventListener("submit", registerAdvisor); id("sale-form").addEventListener("submit", registerSale);
    id("btn-show-register").addEventListener("click", () => { id("auth-view").classList.add("hidden"); id("register-view").classList.remove("hidden"); });
    id("btn-back-login").addEventListener("click", showAuthView); id("btn-logout").addEventListener("click", logout);
    id("btn-menu").addEventListener("click", () => id("sidebar").classList.toggle("open")); id("btn-close-menu").addEventListener("click", closeSidebar);
    id("filtroAsesor").addEventListener("input", renderAdvisorTable);
    ["filtroAdminTexto","filtroAsesorAdmin","filtroEstadoAdmin","filtroTipoAdmin","filtroServicioAdmin","filtroZonaAdmin","filtroDesdeAdmin","filtroHastaAdmin"].forEach(x => { id(x).addEventListener("input", renderAdmin); id(x).addEventListener("change", renderAdmin); });
    id("btn-clear-filters").addEventListener("click", clearAdminFilters); id("btn-preview-report").addEventListener("click", () => previewReport()); id("btn-close-report-preview").addEventListener("click", closeReportPreview); id("btn-print-report").addEventListener("click", () => printReport()); id("btn-pdf-report").addEventListener("click", () => downloadPDF()); id("btn-excel-report").addEventListener("click", downloadExcel);
    id("admin-user-form").addEventListener("submit", saveAdminUser); id("btn-cancel-user-edit").addEventListener("click", resetUserForm);
    id("config-form").addEventListener("submit", saveConfig); id("btn-remove-logo").addEventListener("click", removeLogo);
    id("btn-asesor-report").addEventListener("click", previewAdvisorReport); id("btn-asesor-print").addEventListener("click", printAdvisorReport); id("btn-asesor-pdf").addEventListener("click", downloadAdvisorPDF);
    id("btn-download-backup").addEventListener("click", downloadBackup);
  }

  async function login(e) { e.preventDefault(); const email=value("login-email"), password=id("login-password").value; setButtonBusy(e.submitter,true,"Ingresando..."); const {data,error}=await sbClient.auth.signInWithPassword({email,password}); setButtonBusy(e.submitter,false,"Ingresar"); if(error){showToast(authError(error),true);return;} await initializeSession(data.user); }

  async function registerAdvisor(e) {
    e.preventDefault(); const password=id("reg-password").value, confirm=id("reg-password-confirm").value;
    if(password!==confirm){showToast("Las contraseñas no coinciden.",true);return;} if(password.length<6){showToast("La contraseña debe tener mínimo 6 caracteres.",true);return;}
    const payload={nombre:value("reg-nombre"),apellido:value("reg-apellido"),documento:value("reg-documento"),telefono:value("reg-telefono"),zona:value("reg-zona") ,rol:"asesor"};
    setButtonBusy(e.submitter,true,"Registrando..."); const {data,error}=await sbClient.auth.signUp({email:value("reg-email"),password,options:{data:payload}}); setButtonBusy(e.submitter,false,"Registrar asesor");
    if(error){showToast(authError(error),true);return;} id("register-form").reset(); if(data.session){showToast("Asesor registrado correctamente.");await initializeSession(data.user);}else{showToast("Registro creado. Revisa el correo para confirmar la cuenta.");showAuthView();}
  }

  async function initializeSession(user) {
    currentUser=user;
    const {data:profile,error}=await sbClient.from("perfiles").select("*").eq("id",user.id).single();
    if(error){console.error(error);await sbClient.auth.signOut();showToast("No fue posible cargar tu perfil. Ejecuta el SQL actualizado.",true);return;}
    if(profile.activo === false){await sbClient.auth.signOut();showToast("Tu usuario está inhabilitado. Contacta al administrador.",true);return;}
    currentProfile=profile; await loadConfig(); updateSessionHeader(); buildSidebar();
    if(profile.rol==="administrador"){await loadAdminData();showView("admin-dashboard");} else {await loadAdvisorData();showView("vista-asesor");}
  }

  async function loadConfig(){const {data}=await sbClient.from("configuracion").select("color_principal,logo_url").eq("id",1).maybeSingle(); if(data) config=data; applyTheme(); renderConfig();}

  async function loadAdvisorData(){const {data,error}=await sbClient.from("ventas").select("*").eq("asesor_id",currentUser.id).order("fecha_venta",{ascending:false}).order("id",{ascending:false}); if(error){console.error(error);showToast("No fue posible cargar tus operaciones.",true);return;} sales=data||[]; applyAdvisorProfile();renderAdvisorTable();updateAdvisorDashboard();}

  async function loadAdminData(){
    const [sr,ar]=await Promise.all([
      sbClient.from("ventas").select(`*, perfiles:asesor_id (id,nombre,apellido,zona,email,meta_mensual,activo)`).order("fecha_venta",{ascending:false}).order("id",{ascending:false}),
      sbClient.from("perfiles").select("*").eq("rol","asesor").order("nombre",{ascending:true}).order("apellido",{ascending:true})
    ]);
    if(sr.error){console.error(sr.error);showToast("No fue posible cargar las operaciones.",true);return;} if(ar.error){console.error(ar.error);showToast("No fue posible cargar los asesores.",true);return;}
    sales=sr.data||[]; advisors=ar.data||[]; populateAdminFilters(); renderAdmin(); renderUsers(); updateAdminDashboard(); renderConfig();
  }

  async function registerSale(e){
    e.preventDefault(); if(!currentUser||!currentProfile){showToast("Tu sesión no está disponible.",true);return;}
    const row={asesor_id:currentUser.id,tipo_operacion:value("tipoOperacion"),codigo_cliente:value("codigoCliente"),servicio:value("servicio"),descripcion_servicio:value("descripcionServicio"),zona:value("zona"),fecha_venta:id("fechaVenta").value,estado_instalacion:"PENDIENTE",fecha_instalacion:null};
    if(!row.tipo_operacion||!row.codigo_cliente||!row.servicio||!row.descripcion_servicio||!row.zona||!row.fecha_venta){showToast("Completa todos los campos obligatorios.",true);return;}
    const {data,error}=await sbClient.from("ventas").insert(row).select().single(); if(error){console.error(error);showToast(error.message||"No fue posible registrar la operación.",true);return;}
    e.target.reset();applyAdvisorProfile();setTodayDefault();sales.unshift(data);renderAdvisorTable();updateAdvisorDashboard();showToast(`${row.tipo_operacion} registrada correctamente.`);
  }

  async function setInstallation(id,state){
    if(!currentProfile||currentProfile.rol!=="administrador")return;
    if(state==="CANCELADA"){
      const {data,error}=await sbClient.from("ventas").update({estado_instalacion:"CANCELADA",fecha_instalacion:null}).eq("id",id).select(`*,perfiles:asesor_id (id,nombre,apellido,zona,email,meta_mensual,activo)`).single(); if(error){showToast("No fue posible cancelar la operación.",true);return;} updateSaleLocal(data);showToast("Operación marcada como cancelada.");return;
    }
    const input=document.getElementById(`date-${id}`); if(!input?.value){showToast("Selecciona la fecha de instalación.",true);return;}
    const {data,error}=await sbClient.from("ventas").update({fecha_instalacion:input.value,estado_instalacion:"REALIZADA"}).eq("id",id).select(`*,perfiles:asesor_id (id,nombre,apellido,zona,email,meta_mensual,activo)`).single(); if(error){showToast("No fue posible actualizar la instalación.",true);return;} updateSaleLocal(data);showToast("Instalación marcada como realizada.");
  }
  function updateSaleLocal(data){const i=sales.findIndex(x=>x.id===data.id);if(i>=0)sales[i]=data;renderAdmin();updateAdminDashboard();}

  async function deleteSale(id){if(!confirm("¿Eliminar definitivamente esta operación? Esta acción no se puede deshacer."))return;const {error}=await sbClient.from("ventas").delete().eq("id",id);if(error){showToast("No fue posible eliminar la venta. Verifica las políticas RLS.",true);return;}sales=sales.filter(x=>x.id!==id);renderAdmin();updateAdminDashboard();showToast("Operación eliminada.");}

  function buildSidebar(){const nav=id("sidebar-nav");const admin=currentProfile?.rol==="administrador";const items=admin?[ ["admin-dashboard","▦","Dashboard"],["vista-admin","▤","Operaciones"],["vista-usuarios","♙","Usuarios"],["vista-configuracion","⚙","Configuración"],["vista-respaldo","⭳","Respaldo"] ]:[["vista-asesor","▦","Mi dashboard"],["vista-asesor","＋","Registrar operación"],["vista-asesor","▤","Mis operaciones"]];nav.innerHTML=items.map(([target,icon,label])=>`<button class="nav-item" type="button" data-target="${target}" data-anchor="${target==='vista-asesor'?label:''}"><span>${icon}</span>${label}</button>`).join("");nav.querySelectorAll(".nav-item").forEach(b=>b.addEventListener("click",()=>{showView(b.dataset.target);if(b.dataset.anchor==="Registrar operación")id("asesor-form-section").scrollIntoView({behavior:"smooth"});if(b.dataset.anchor==="Mis operaciones")document.querySelector("#vista-asesor .table-card").scrollIntoView({behavior:"smooth"});closeSidebar();}));}
  function closeSidebar(){id("sidebar").classList.remove("open");}

  function updateSessionHeader(){const name=[currentProfile?.nombre,currentProfile?.apellido].filter(Boolean).join(" ")||"Usuario", role=currentProfile?.rol==="administrador"?"Administrador":"Asesor";id("user-name").textContent=name;id("user-role").textContent=role;id("user-avatar").textContent=name.charAt(0).toUpperCase();id("sidebar-user-name").textContent=name;id("sidebar-user-role").textContent=role;id("session-area").classList.remove("hidden");id("btn-menu").classList.remove("hidden");id("sidebar").classList.remove("hidden");}
  function applyAdvisorProfile(){const zona=currentProfile?.zona||"";id("zona").value=zona;id("asesor-zone-badge").textContent=`Zona: ${zona||"Sin asignar"}`;id("asesor-welcome").textContent=`Registra operaciones y consulta tu avance. Zona asignada: ${zona||"sin asignar"}.`;}

  function renderAdvisorTable(){const tabla=id("tabla-asesor"),filtro=value("filtroAsesor").toLowerCase(),filtered=sales.filter(s=>[s.tipo_operacion,s.codigo_cliente,s.servicio,s.descripcion_servicio,s.zona,s.fecha_venta,s.estado_instalacion].join(" ").toLowerCase().includes(filtro));tabla.innerHTML=filtered.length?filtered.map(s=>`<tr><td>#${s.id}</td><td>${operationBadge(s.tipo_operacion)}</td><td>${escapeHTML(s.codigo_cliente)}</td><td>${serviceBadge(s.servicio)}</td><td>${escapeHTML(s.descripcion_servicio)}</td><td>${escapeHTML(s.zona)}</td><td>${formatDate(s.fecha_venta)}</td><td>${installationStatus(s.estado_instalacion)}</td><td>${formatDate(s.fecha_instalacion)}</td></tr>`).join(""):`<tr class="empty-row"><td colspan="9">${sales.length?"No se encontraron operaciones.":"No hay operaciones registradas."}</td></tr>`;updateAdvisorStats();}
  function updateAdvisorStats(){const total=sales.length,ventas=sales.filter(s=>s.tipo_operacion==="Venta").length,recon=sales.filter(s=>s.tipo_operacion==="Reconexión").length,done=sales.filter(s=>s.estado_instalacion==="REALIZADA").length;id("asesor-total-count").textContent=total;id("asesor-ventas-count").textContent=ventas;id("asesor-reconexion-count").textContent=recon;id("asesor-complete-count").textContent=done;}
  function updateAdvisorDashboard(){updateAdvisorStats();const now=new Date(),ym=`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`,monthly=sales.filter(s=>s.fecha_venta?.startsWith(ym)&&s.tipo_operacion==="Venta").length,goal=Math.max(1,Number(currentProfile?.meta_mensual)||50),pct=Math.min(100,Math.round(monthly/goal*100));id("asesor-goal-title").textContent=`${monthly} / ${goal} ventas`;id("asesor-goal-detail").textContent=`Meta configurada por administración para ${now.toLocaleDateString("es-CO",{month:"long",year:"numeric"})}.`;id("asesor-goal-bar").style.width=`${pct}%`;id("asesor-goal-percent").textContent=`${pct}%`;}

  function renderAdmin(){const tabla=id("tabla-admin");if(!tabla)return;const filtered=getFilteredAdminSales();tabla.innerHTML=filtered.length?filtered.map(s=>{const a=s.perfiles||{};const name=[a.nombre,a.apellido].filter(Boolean).join(" ")||"—";return `<tr><td>#${s.id}</td><td>${escapeHTML(name)}</td><td>${operationBadge(s.tipo_operacion)}</td><td>${escapeHTML(s.codigo_cliente)}</td><td>${serviceBadge(s.servicio)}</td><td>${escapeHTML(s.descripcion_servicio)}</td><td>${escapeHTML(s.zona)}</td><td>${formatDate(s.fecha_venta)}</td><td>${installationStatus(s.estado_instalacion)}</td><td><input class="installation-date" type="date" id="date-${s.id}" value="${s.fecha_instalacion||""}" ${s.estado_instalacion!=="PENDIENTE"?"disabled":""}></td><td class="action-cell"><button class="btn-save-installation" ${s.estado_instalacion!=="PENDIENTE"?"disabled":""} onclick="setInstallation(${s.id},'REALIZADA')">Realizar</button><button class="btn-cancel-sale" ${s.estado_instalacion!=="PENDIENTE"?"disabled":""} onclick="setInstallation(${s.id},'CANCELADA')">Cancelar</button><button class="btn-delete" onclick="deleteSale(${s.id})">Eliminar</button></td></tr>`;}).join(""):`<tr class="empty-row"><td colspan="11">${sales.length?"No se encontraron operaciones con los filtros seleccionados.":"No hay operaciones registradas."}</td></tr>`;id("admin-result-count").textContent=`${filtered.length} resultado${filtered.length===1?"":"s"}`;}
  function getFilteredAdminSales(){const text=value("filtroAdminTexto").toLowerCase(),advisor=id("filtroAsesorAdmin").value,state=id("filtroEstadoAdmin").value,type=id("filtroTipoAdmin").value,service=id("filtroServicioAdmin").value,zone=id("filtroZonaAdmin").value,from=id("filtroDesdeAdmin").value,to=id("filtroHastaAdmin").value;return sales.filter(s=>{const a=s.perfiles||{},search=[a.nombre,a.apellido,a.email,s.tipo_operacion,s.codigo_cliente,s.servicio,s.descripcion_servicio,s.zona,s.fecha_venta].join(" ").toLowerCase();return(!text||search.includes(text))&&(!advisor||s.asesor_id===advisor)&&(!state||s.estado_instalacion===state)&&(!type||s.tipo_operacion===type)&&(!service||s.servicio===service)&&(!zone||s.zona===zone)&&(!from||s.fecha_venta>=from)&&(!to||s.fecha_venta<=to);});}
  function populateAdminFilters(){const as=id("filtroAsesorAdmin"),zone=id("filtroZonaAdmin"),aVal=as.value,zVal=zone.value;as.innerHTML='<option value="">Todos los asesores</option>'+advisors.map(a=>`<option value="${a.id}">${escapeHTML([a.nombre,a.apellido].filter(Boolean).join(" ")||a.email)}</option>`).join("");as.value=aVal;const zones=[...new Set(sales.map(s=>s.zona).filter(Boolean))].sort((a,b)=>a.localeCompare(b));zone.innerHTML='<option value="">Todas las zonas</option>'+zones.map(z=>`<option>${escapeHTML(z)}</option>`).join("");zone.value=zVal;}

  function updateAdminDashboard(){const total=sales.length,ventas=sales.filter(s=>s.tipo_operacion==="Venta").length,p=sales.filter(s=>s.estado_instalacion==="PENDIENTE").length,r=sales.filter(s=>s.estado_instalacion==="REALIZADA").length,c=sales.filter(s=>s.estado_instalacion==="CANCELADA").length;setText("dash-total",total);setText("dash-ventas",ventas);setText("dash-pendientes",p);setText("dash-realizadas",r);setText("dash-canceladas",c);const now=new Date(),ym=`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`,monthly=sales.filter(s=>s.fecha_venta?.startsWith(ym)&&s.tipo_operacion==="Venta");id("dash-goals-list").innerHTML=advisors.map(a=>{const n=[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email,count=monthly.filter(s=>s.asesor_id===a.id).length,g=Math.max(1,Number(a.meta_mensual)||50),pct=Math.min(100,Math.round(count/g*100));return `<div class="goal-list-row"><div><strong>${escapeHTML(n)}</strong><small>${count} / ${g} ventas</small></div><div class="mini-progress"><span style="width:${pct}%"></span></div><b>${pct}%</b></div>`;}).join("")||'<p class="muted">No hay asesores registrados.</p>';const counts=SERVICES.map(s=>({s,n:monthly.filter(x=>x.servicio===s).length}));const max=Math.max(1,...counts.map(x=>x.n));id("dash-services-list").innerHTML=counts.map(x=>`<div class="mini-bar-row"><span>${x.s}</span><div><i style="width:${x.n/max*100}%"></i></div><strong>${x.n}</strong></div>`).join("");}

  function renderUsers(){const tbody=id("tabla-usuarios");if(!tbody)return;tbody.innerHTML=advisors.map(a=>{const name=[a.nombre,a.apellido].filter(Boolean).join(" ")||"—";return `<tr><td><strong>${escapeHTML(name)}</strong></td><td>${escapeHTML(a.email||"—")}</td><td>${escapeHTML(a.zona||"—")}</td><td>${Number(a.meta_mensual)||50}</td><td>${a.activo===false?'<span class="badge badge-disabled">Inhabilitado</span>':'<span class="badge badge-active">Activo</span>'}</td><td class="action-cell"><button class="btn-small" onclick="editAdvisor('${a.id}')">Editar</button><button class="btn-small" onclick="toggleAdvisor('${a.id}',${a.activo!==false})">${a.activo===false?"Habilitar":"Inhabilitar"}</button><button class="btn-delete" onclick="deleteAdvisor('${a.id}')">Eliminar</button></td></tr>`;}).join("")||'<tr class="empty-row"><td colspan="6">No hay asesores registrados.</td></tr>';}

  async function saveAdminUser(e){e.preventDefault();const idUser=id("admin-user-id").value;const body={nombre:value("admin-user-nombre"),apellido:value("admin-user-apellido"),documento:value("admin-user-documento"),telefono:value("admin-user-telefono"),zona:value("admin-user-zona"),email:value("admin-user-email"),meta_mensual:Math.max(1,Number(id("admin-user-meta").value)||50)};if(!idUser){const password=id("admin-user-password").value;if(password.length<6){showToast("La contraseña debe tener mínimo 6 caracteres.",true);return;}const session=(await sbClient.auth.getSession()).data.session;const {data,error}=await fetchAdminFunction("create",{...body,password});if(error){showToast(error,true);return;}showToast("Asesor creado correctamente.");resetUserForm();await loadAdminData();return;}const result=await fetchAdminFunction("update",{user_id:idUser,...body});if(result.error){showToast(result.error,true);return;}showToast("Asesor actualizado.");resetUserForm();await loadAdminData();}
  async function fetchAdminFunction(action,payload){const {data:{session}}=await sbClient.auth.getSession();if(!session)return{error:"Sesión no disponible."};try{const r=await fetch(`${SUPABASE_URL}/functions/v1/admin-users`,{method:"POST",headers:{"Content-Type":"application/json",Authorization:`Bearer ${session.access_token}`},body:JSON.stringify({action,...payload})});const j=await r.json().catch(()=>({}));return r.ok?{data:j}:{error:j.error||`Error ${r.status}`};}catch(e){return{error:"No se pudo contactar la función de administración. Debes desplegar supabase/functions/admin-users."};}}
  function editAdvisor(uid){const a=advisors.find(x=>x.id===uid);if(!a)return;id("admin-user-id").value=a.id;["nombre","apellido","documento","telefono","zona","email"].forEach(k=>id(`admin-user-${k}`).value=a[k]||"");id("admin-user-meta").value=Number(a.meta_mensual)||50;id("admin-user-password").value="";id("btn-save-user").textContent="Actualizar asesor";id("btn-cancel-user-edit").classList.remove("hidden");document.getElementById("vista-usuarios").scrollIntoView({behavior:"smooth"});}
  async function toggleAdvisor(uid,active){const {error}=await sbClient.from("perfiles").update({activo:!active}).eq("id",uid);if(error){showToast(error.message,true);return;}showToast(active?"Asesor inhabilitado.":"Asesor habilitado.");await loadAdminData();}
  async function deleteAdvisor(uid){const a=advisors.find(x=>x.id===uid);if(!a)return;if(!confirm(`¿Eliminar a ${[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email}? Solo se podrá eliminar si no tiene operaciones registradas.`))return;const result=await fetchAdminFunction("delete",{user_id:uid});if(result.error){showToast(result.error,true);return;}showToast("Asesor eliminado.");await loadAdminData();}
  function resetUserForm(){id("admin-user-form").reset();id("admin-user-id").value="";id("admin-user-meta").value=50;id("btn-save-user").textContent="Crear asesor";id("btn-cancel-user-edit").classList.add("hidden");}

  async function saveConfig(e){e.preventDefault();let logo=config.logo_url||"";const file=id("config-logo").files[0];if(file){if(file.size>2*1024*1024){showToast("La imagen debe pesar máximo 2 MB.",true);return;}logo=await fileToDataURL(file);}const color=id("config-color").value;const {error}=await sbClient.from("configuracion").upsert({id:1,color_principal:color,logo_url:logo,updated_by:currentUser.id},{onConflict:"id"});if(error){showToast(error.message,true);return;}config={color_principal:color,logo_url:logo};applyTheme();renderConfig();showToast("Configuración guardada.");}
  async function removeLogo(){const {error}=await sbClient.from("configuracion").upsert({id:1,color_principal:config.color_principal,logo_url:"",updated_by:currentUser.id},{onConflict:"id"});if(error){showToast(error.message,true);return;}config.logo_url="";renderConfig();showToast("Imagen retirada del reporte.");}
  function renderConfig(){id("config-color").value=config.color_principal||"#8b5cf6";id("logo-preview").innerHTML=config.logo_url?`<img src="${config.logo_url}" alt="Logo de empresa">`:'<span>LOGO</span>';}
  function applyTheme(){document.documentElement.style.setProperty("--purple-primary",config.color_principal||"#8b5cf6");}

  function clearAdminFilters(){["filtroAdminTexto","filtroAsesorAdmin","filtroEstadoAdmin","filtroTipoAdmin","filtroServicioAdmin","filtroZonaAdmin","filtroDesdeAdmin","filtroHastaAdmin"].forEach(x=>id(x).value="");renderAdmin();}

  function buildReportHTML(){const filtered=getFilteredAdminSales(),total=filtered.length,ventas=filtered.filter(s=>s.tipo_operacion==="Venta").length,recon=filtered.filter(s=>s.tipo_operacion==="Reconexión").length,otros=filtered.filter(s=>s.tipo_operacion==="Otros").length,real=filtered.filter(s=>s.estado_instalacion==="REALIZADA").length,pending=filtered.filter(s=>s.estado_instalacion==="PENDIENTE").length,cancel=filtered.filter(s=>s.estado_instalacion==="CANCELADA").length,pct=n=>total?Math.round(n/total*100):0;
    const advisorMap={};filtered.forEach(s=>{const a=s.perfiles||{},n=[a.nombre,a.apellido].filter(Boolean).join(" ")||"Sin asesor";if(!advisorMap[n])advisorMap[n]={ventas:0,meta:Math.max(1,Number(a.meta_mensual)||50)};if(s.tipo_operacion==="Venta")advisorMap[n].ventas++;});const advisorRows=Object.entries(advisorMap).sort((a,b)=>b[1].ventas-a[1].ventas).map(([n,d])=>{const gp=Math.min(100,Math.round(d.ventas/d.meta*100));return `<div class="print-advisor-row"><div class="print-advisor-label"><span>${escapeHTML(n)}</span><strong>${d.ventas}/${d.meta} ventas · ${gp}%</strong></div><div class="print-bar-track"><div class="print-bar-fill" style="width:${gp}%"></div></div></div>`;}).join("")||'<div class="print-empty-chart">Sin datos</div>';
    const desde=value("filtroDesdeAdmin"),hasta=value("filtroHastaAdmin"),period=desde||hasta?`${desde?formatDate(desde):"Inicio"} – ${hasta?formatDate(hasta):"Actual"}`:"Todos los periodos";const rows=filtered.map(s=>{const a=s.perfiles||{},n=[a.nombre,a.apellido].filter(Boolean).join(" ")||"—";return `<tr><td>${escapeHTML(n)}</td><td>${escapeHTML(s.tipo_operacion||"—")}</td><td>${escapeHTML([s.servicio,s.descripcion_servicio].filter(Boolean).join(" · "))}</td><td>${escapeHTML(s.zona||"—")}</td><td>${escapeHTML(statusLabel(s.estado_instalacion))}</td></tr>`}).join("");
    return `<div class="print-report-sheet">${config.logo_url?`<div class="print-logo"><img src="${config.logo_url}" alt="Logo"></div>`:""}<div class="print-header"><div><span class="print-kicker">REPORTE DE OPERACIONES</span><h1>Ventas e instalaciones</h1><p>Periodo: <strong>${escapeHTML(period)}</strong></p></div><div class="print-generated">Generado: ${new Date().toLocaleString("es-CO")}</div></div><div class="print-summary"><div class="print-summary-card"><span>Total operaciones</span><strong>${total}</strong></div><div class="print-summary-card"><span>Realizadas</span><strong>${real}</strong></div><div class="print-summary-card"><span>Pendientes</span><strong>${pending}</strong></div><div class="print-summary-card"><span>Canceladas</span><strong>${cancel}</strong></div></div><section class="print-charts"><div class="print-chart-card"><h2>Operaciones</h2><div class="print-donut" style="--first:${pct(ventas)*3.6}deg"><div class="print-donut-center"><strong>${total}</strong><span>total</span></div></div><div class="print-legend"><span>Venta <strong>${pct(ventas)}%</strong></span><span>Reconexión <strong>${pct(recon)}%</strong></span><span>Otros <strong>${pct(otros)}%</strong></span></div></div><div class="print-chart-card"><h2>Estado</h2><div class="print-donut" style="--first:${pct(real)*3.6}deg"><div class="print-donut-center"><strong>${pct(real)}%</strong><span>realizadas</span></div></div><div class="print-legend"><span>Realizada <strong>${pct(real)}%</strong></span><span>Pendiente <strong>${pct(pending)}%</strong></span><span>Cancelada <strong>${pct(cancel)}%</strong></span></div></div><div class="print-chart-card print-advisor-chart"><h2>Cumplimiento de meta por asesor</h2>${advisorRows}</div></section><section class="print-table-section"><div class="print-table-title"><div><span class="print-kicker">DETALLE</span><h2>Operaciones registradas</h2></div><strong>${total} resultado${total===1?"":"s"}</strong></div><div class="print-table-scroll"><table><thead><tr><th>Asesor</th><th>Operación</th><th>Servicio</th><th>Zona</th><th>Estado</th></tr></thead><tbody>${rows||'<tr><td colspan="5" class="print-empty-row">No hay registros.</td></tr>'}</tbody></table></div></section></div>`;
  }
  function previewReport(builder=buildReportHTML){const modal=id("report-preview-modal"),content=id("report-preview-content");if(!modal||!content)return;content.innerHTML=builder();modal.classList.remove("hidden");modal.setAttribute("aria-hidden","false");document.body.classList.add("report-preview-open");}
  function closeReportPreview(){const modal=id("report-preview-modal");if(!modal)return;modal.classList.add("hidden");modal.setAttribute("aria-hidden","true");document.body.classList.remove("report-preview-open");}
  function printReport(builder=buildReportHTML){id("print-report").innerHTML=builder();window.print();}
  async function downloadPDF(builder=buildReportHTML,filePrefix="reporte-ventas"){const area=id("print-report");area.innerHTML=builder();area.classList.add("pdf-rendering");try{const canvas=await html2canvas(area,{scale:2,useCORS:true,backgroundColor:"#ffffff"});const {jsPDF}=window.jspdf;const pdf=new jsPDF({orientation:"landscape",unit:"mm",format:"a4"});const pageW=297,pageH=210,margin=8,imgW=pageW-margin*2,imgH=canvas.height*imgW/canvas.width;let y=margin;let sourceY=0;const pxPerPage=canvas.width*(pageH-margin*2)/imgW;while(sourceY<canvas.height){const h=Math.min(pxPerPage,canvas.height-sourceY);const pageCanvas=document.createElement("canvas");pageCanvas.width=canvas.width;pageCanvas.height=h;pageCanvas.getContext("2d").drawImage(canvas,0,sourceY,canvas.width,h,0,0,canvas.width,h);const pageImg=pageCanvas.toDataURL("image/jpeg",0.95);const hMm=h*imgW/canvas.width;if(sourceY>0)pdf.addPage();pdf.addImage(pageImg,"JPEG",margin,margin,imgW,hMm);sourceY+=h;}const now=new Date().toISOString().slice(0,10);pdf.save(`${filePrefix}-${now}.pdf`);showToast("PDF descargado correctamente.");}catch(e){console.error(e);showToast("No fue posible generar el PDF.",true);}finally{area.classList.remove("pdf-rendering");}}

  // Reporte de avance individual para el asesor (usa "sales", que ya viene filtrado a sus propias operaciones).
  function buildAdvisorReportHTML(){
    const list=sales,total=list.length,ventas=list.filter(s=>s.tipo_operacion==="Venta").length,recon=list.filter(s=>s.tipo_operacion==="Reconexión").length,otros=list.filter(s=>s.tipo_operacion==="Otros").length,real=list.filter(s=>s.estado_instalacion==="REALIZADA").length,pending=list.filter(s=>s.estado_instalacion==="PENDIENTE").length,cancel=list.filter(s=>s.estado_instalacion==="CANCELADA").length,pct=n=>total?Math.round(n/total*100):0;
    const now=new Date(),ym=`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`,monthly=list.filter(s=>s.fecha_venta?.startsWith(ym)&&s.tipo_operacion==="Venta").length,goal=Math.max(1,Number(currentProfile?.meta_mensual)||50),gp=Math.min(100,Math.round(monthly/goal*100));
    const nombre=[currentProfile?.nombre,currentProfile?.apellido].filter(Boolean).join(" ")||currentProfile?.email||"Asesor";
    const rows=list.map(s=>`<tr><td>${operationBadge(s.tipo_operacion)}</td><td>${escapeHTML([s.servicio,s.descripcion_servicio].filter(Boolean).join(" · "))}</td><td>${escapeHTML(s.zona||"—")}</td><td>${escapeHTML(formatDate(s.fecha_venta))}</td><td>${escapeHTML(statusLabel(s.estado_instalacion))}</td></tr>`).join("");
    return `<div class="print-report-sheet">${config.logo_url?`<div class="print-logo"><img src="${config.logo_url}" alt="Logo"></div>`:""}<div class="print-header"><div><span class="print-kicker">REPORTE DE AVANCE</span><h1>${escapeHTML(nombre)}</h1><p>Zona: <strong>${escapeHTML(currentProfile?.zona||"—")}</strong></p></div><div class="print-generated">Generado: ${new Date().toLocaleString("es-CO")}</div></div><div class="print-summary"><div class="print-summary-card"><span>Total operaciones</span><strong>${total}</strong></div><div class="print-summary-card"><span>Realizadas</span><strong>${real}</strong></div><div class="print-summary-card"><span>Pendientes</span><strong>${pending}</strong></div><div class="print-summary-card"><span>Canceladas</span><strong>${cancel}</strong></div></div><section class="print-charts"><div class="print-chart-card"><h2>Operaciones</h2><div class="print-donut" style="--first:${pct(ventas)*3.6}deg"><div class="print-donut-center"><strong>${total}</strong><span>total</span></div></div><div class="print-legend"><span>Venta <strong>${pct(ventas)}%</strong></span><span>Reconexión <strong>${pct(recon)}%</strong></span><span>Otros <strong>${pct(otros)}%</strong></span></div></div><div class="print-chart-card"><h2>Estado</h2><div class="print-donut" style="--first:${pct(real)*3.6}deg"><div class="print-donut-center"><strong>${pct(real)}%</strong><span>realizadas</span></div></div><div class="print-legend"><span>Realizada <strong>${pct(real)}%</strong></span><span>Pendiente <strong>${pct(pending)}%</strong></span><span>Cancelada <strong>${pct(cancel)}%</strong></span></div></div><div class="print-chart-card print-advisor-chart"><h2>Meta mensual</h2><div class="print-advisor-row"><div class="print-advisor-label"><span>${escapeHTML(now.toLocaleDateString("es-CO",{month:"long",year:"numeric"}))}</span><strong>${monthly}/${goal} ventas · ${gp}%</strong></div><div class="print-bar-track"><div class="print-bar-fill" style="width:${gp}%"></div></div></div></div></section><section class="print-table-section"><div class="print-table-title"><div><span class="print-kicker">DETALLE</span><h2>Mis operaciones</h2></div><strong>${total} resultado${total===1?"":"s"}</strong></div><div class="print-table-scroll"><table><thead><tr><th>Operación</th><th>Servicio</th><th>Zona</th><th>Fecha</th><th>Estado</th></tr></thead><tbody>${rows||'<tr><td colspan="5" class="print-empty-row">No hay registros.</td></tr>'}</tbody></table></div></section></div>`;
  }
  function previewAdvisorReport(){previewReport(buildAdvisorReportHTML);}
  function printAdvisorReport(){printReport(buildAdvisorReportHTML);}
  function downloadAdvisorPDF(){downloadPDF(buildAdvisorReportHTML,"mi-reporte");}

  function downloadExcel(){
    try{
      if(!window.XLSX){showToast("No se pudo cargar el módulo de Excel.",true);return;}
      const filtered=getFilteredAdminSales();
      const total=filtered.length, ventas=filtered.filter(s=>s.tipo_operacion==="Venta").length, recon=filtered.filter(s=>s.tipo_operacion==="Reconexión").length, otros=filtered.filter(s=>s.tipo_operacion==="Otros").length, real=filtered.filter(s=>s.estado_instalacion==="REALIZADA").length, pending=filtered.filter(s=>s.estado_instalacion==="PENDIENTE").length, cancel=filtered.filter(s=>s.estado_instalacion==="CANCELADA").length;
      const pct=n=>total?Math.round(n/total*100):0;
      const advisorMap={};filtered.forEach(s=>{const a=s.perfiles||{},n=[a.nombre,a.apellido].filter(Boolean).join(" ")||"Sin asesor";if(!advisorMap[n])advisorMap[n]={ventas:0,meta:Math.max(1,Number(a.meta_mensual)||50)};if(s.tipo_operacion==="Venta")advisorMap[n].ventas++;});
      const detail=filtered.map(s=>{const a=s.perfiles||{};return {"Asesor":[a.nombre,a.apellido].filter(Boolean).join(" ")||"—","Operación":s.tipo_operacion||"—","Servicio":[s.servicio,s.descripcion_servicio].filter(Boolean).join(" · "),"Zona":s.zona||"—","Estado":statusLabel(s.estado_instalacion)}});
      const ws=window.XLSX.utils.json_to_sheet(detail.length?detail:[{"Asesor":"","Operación":"","Servicio":"","Zona":"","Estado":""}],{header:["Asesor","Operación","Servicio","Zona","Estado"]});
      ws["!cols"]=[{wch:25},{wch:16},{wch:48},{wch:20},{wch:16}];
      const summary=[
        ["REPORTE DE OPERACIONES"], ["Periodo", value("filtroDesdeAdmin")||value("filtroHastaAdmin")?`${value("filtroDesdeAdmin")?formatDate(value("filtroDesdeAdmin")):"Inicio"} – ${value("filtroHastaAdmin")?formatDate(value("filtroHastaAdmin")):"Actual"}`:"Todos los periodos"], [],
        ["RESUMEN GENERAL"],["Indicador","Cantidad","Porcentaje"],["Total operaciones",total,"100%"],["Ventas",ventas,`${pct(ventas)}%`],["Reconexiones",recon,`${pct(recon)}%`],["Otros",otros,`${pct(otros)}%`],["Realizadas",real,`${pct(real)}%`],["Pendientes",pending,`${pct(pending)}%`],["Canceladas",cancel,`${pct(cancel)}%`],[],
        ["GRÁFICO · OPERACIONES"],["Categoría","Cantidad","%","Representación"],["Venta",ventas,pct(ventas),"█".repeat(Math.max(0,Math.round(pct(ventas)/5)))],["Reconexión",recon,pct(recon),"█".repeat(Math.max(0,Math.round(pct(recon)/5)))],["Otros",otros,pct(otros),"█".repeat(Math.max(0,Math.round(pct(otros)/5)))],[],
        ["GRÁFICO · ESTADO"],["Estado","Cantidad","%","Representación"],["Realizada",real,pct(real),"█".repeat(Math.max(0,Math.round(pct(real)/5)))],["Pendiente",pending,pct(pending),"█".repeat(Math.max(0,Math.round(pct(pending)/5)))],["Cancelada",cancel,pct(cancel),"█".repeat(Math.max(0,Math.round(pct(cancel)/5)))],[],
        ["GRÁFICO · CUMPLIMIENTO DE META POR ASESOR"],["Asesor","Ventas","Meta","% cumplimiento","Representación"],
        ...Object.entries(advisorMap).sort((a,b)=>b[1].ventas-a[1].ventas).map(([n,d])=>{const gp=Math.min(100,Math.round(d.ventas/d.meta*100));return [n,d.ventas,d.meta,gp,"█".repeat(Math.max(0,Math.round(gp/5)))];})
      ];
      const wr=window.XLSX.utils.aoa_to_sheet(summary);wr["!cols"]=[{wch:34},{wch:15},{wch:15},{wch:28}];
      const wb=window.XLSX.utils.book_new();window.XLSX.utils.book_append_sheet(wb,wr,"Resumen y gráficos");window.XLSX.utils.book_append_sheet(wb,ws,"Detalle");
      window.XLSX.writeFile(wb,`reporte-ventas-${new Date().toISOString().slice(0,10)}.xlsx`);showToast("Excel descargado con resumen y gráficos.");
    }catch(e){console.error(e);showToast("No fue posible generar el Excel.",true);}
  }

  function showAuthView(){["auth-view","register-view","vista-asesor","admin-dashboard","vista-admin","vista-usuarios","vista-configuracion"].forEach(x=>id(x).classList.add("hidden"));id("auth-view").classList.remove("hidden");id("session-area").classList.add("hidden");id("btn-menu").classList.add("hidden");id("sidebar").classList.add("hidden");}
  function showView(viewId){["auth-view","register-view","vista-asesor","admin-dashboard","vista-admin","vista-usuarios","vista-configuracion","vista-respaldo"].forEach(x=>id(x).classList.add("hidden"));id(viewId).classList.remove("hidden");if(viewId!=="auth-view"&&currentProfile){id("session-area").classList.remove("hidden");id("btn-menu").classList.remove("hidden");id("sidebar").classList.remove("hidden");}}
  async function logout(){const {error}=await sbClient.auth.signOut();if(error)showToast("No fue posible cerrar la sesión.",true);}
  function installationStatus(s){if(s==="REALIZADA")return '<span class="badge badge-complete">Realizada</span>';if(s==="CANCELADA")return '<span class="badge badge-cancelled">Cancelada</span>';return '<span class="badge badge-pending">Pendiente</span>';}
  function statusLabel(s){return s==="REALIZADA"?"Realizada":s==="CANCELADA"?"Cancelada":"Pendiente";}
  function operationBadge(t){if(t==="Reconexión")return '<span class="badge badge-reconnection">Reconexión</span>';if(t==="Otros")return '<span class="badge badge-other">Otros</span>';return '<span class="badge badge-sale">Venta</span>';}
  function serviceBadge(s){return `<span class="badge badge-service">${escapeHTML(s||"Otros")}</span>`;}
  function formatDate(d){if(!d)return "—";const p=d.split("-");return p.length===3?`${p[2]}/${p[1]}/${p[0]}`:escapeHTML(d);}
  function setTodayDefault(){const x=id("fechaVenta");if(x&&!x.value)x.value=getTodayISO();}function getTodayISO(){const n=new Date(),o=n.getTimezoneOffset(),l=new Date(n.getTime()-o*60000);return l.toISOString().slice(0,10);}
  function value(x){return id(x).value.trim();}function id(x){return document.getElementById(x);}function setText(x,v){if(id(x))id(x).textContent=v;}
  function escapeHTML(v){return String(v??"").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;");}
  async function downloadBackup(){
    const btn=id("btn-download-backup"),status=id("backup-status");
    if(!window.XLSX){showToast("No se pudo cargar el módulo de Excel.",true);return;}
    setButtonBusy(btn,true,"Generando respaldo...");
    try{
      const [ventasRes,perfilesRes]=await Promise.all([
        sbClient.from("ventas").select("*").order("id",{ascending:true}),
        sbClient.from("perfiles").select("*").order("created_at",{ascending:true})
      ]);
      if(ventasRes.error||perfilesRes.error){console.error(ventasRes.error||perfilesRes.error);showToast("No fue posible generar el respaldo.",true);return;}
      const wb=window.XLSX.utils.book_new();
      const wsVentas=window.XLSX.utils.json_to_sheet(ventasRes.data&&ventasRes.data.length?ventasRes.data:[{id:""}]);
      const wsPerfiles=window.XLSX.utils.json_to_sheet(perfilesRes.data&&perfilesRes.data.length?perfilesRes.data:[{id:""}]);
      window.XLSX.utils.book_append_sheet(wb,wsVentas,"Ventas");
      window.XLSX.utils.book_append_sheet(wb,wsPerfiles,"Perfiles");
      const now=new Date();
      window.XLSX.writeFile(wb,`respaldo-cabletelco-${now.toISOString().slice(0,10)}.xlsx`);
      status.textContent=`Último respaldo generado: ${now.toLocaleString("es-CO")} · ${ventasRes.data.length} ventas, ${perfilesRes.data.length} perfiles.`;
      showToast("Respaldo generado correctamente.");
    }catch(e){console.error(e);showToast("No fue posible generar el respaldo.",true);}
    finally{setButtonBusy(btn,false,"⭳ Descargar respaldo completo");}
  }

  function setButtonBusy(b,busy,text){if(!b)return;b.disabled=busy;b.textContent=text;}function authError(e){const m=(e?.message||"").toLowerCase();if(m.includes("invalid login credentials"))return "Correo o contraseña incorrectos.";if(m.includes("email not confirmed"))return "Debes confirmar tu correo antes de iniciar sesión.";if(m.includes("user already registered"))return "Ese correo ya está registrado.";return e?.message||"No fue posible completar la operación.";}
  function fileToDataURL(file){return new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=rej;r.readAsDataURL(file);});}
  let toastTimer;function showToast(msg,error=false){const t=id("toast");t.textContent=msg;t.classList.toggle("error",error);t.classList.add("show");clearTimeout(toastTimer);toastTimer=setTimeout(()=>t.classList.remove("show"),3500);}

  window.setInstallation=setInstallation;window.deleteSale=deleteSale;window.editAdvisor=editAdvisor;window.toggleAdvisor=toggleAdvisor;window.deleteAdvisor=deleteAdvisor;
})();
