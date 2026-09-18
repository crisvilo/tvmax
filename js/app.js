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
  const SURVEY_QUESTIONS = {
    q2_servicio: "¿CÓMO CALIFICA EL SERVICIO PRESTADO POR GRUPO TV MAX?",
    q3_tecnica: "¿CÓMO CALIFICA LA ATENCIÓN PRESTADA POR PARTE DEL ÁREA TÉCNICA DE GRUPO TV MAX AL ACERCARSE A SU RESIDENCIA?",
    q4_administrativa: "¿CÓMO CALIFICA LA ATENCIÓN PRESTADA POR PARTE DEL ÁREA ADMINISTRATIVA DE GRUPO TV MAX CUANDO SE ACERCA A LA OFICINA O POR MEDIO DE LLAMADAS?",
    q5_agilidad: "¿CUANDO HA TENIDO DAÑOS O AVERÍAS EN ALGUNO DE LOS SERVICIOS CONTRATADOS, QUÉ TAN ÁGILES Y OPORTUNOS HEMOS SIDO?",
    q6_recomendaria: "RECOMENDARÍA NUESTROS SERVICIOS"
  };
  function isGoalOperation(s){return s.tipo_operacion==="Venta"||s.tipo_operacion==="Reconexión";}
  function excelColRef(n){let s="",m=n+1;while(m>0){const r=(m-1)%26;s=String.fromCharCode(65+r)+s;m=Math.floor((m-1)/26);}return s;}
  function excelRangeRef(sheetName,row0,row1,col){return `'${sheetName.replace(/'/g,"''")}'!$${excelColRef(col)}$${row0+1}:$${excelColRef(col)}$${row1+1}`;}
  function chartPartXml({type,title,seriesName,catRef,catCache,valRef,valCache}){
    const n=catCache.length;
    const catPts=catCache.map((v,i)=>`<c:pt idx="${i}"><c:v>${escapeHTML(v)}</c:v></c:pt>`).join("");
    const valPts=valCache.map((v,i)=>`<c:pt idx="${i}"><c:v>${Number(v)||0}</c:v></c:pt>`).join("");
    const catBlock=`<c:cat><c:strRef><c:f>${catRef}</c:f><c:strCache><c:ptCount val="${n}"/>${catPts}</c:strCache></c:strRef></c:cat>`;
    const valBlock=`<c:val><c:numRef><c:f>${valRef}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${n}"/>${valPts}</c:numCache></c:numRef></c:val>`;
    const ser=`<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>${escapeHTML(seriesName)}</c:v></c:tx>${catBlock}${valBlock}</c:ser>`;
    const body=type==="pie"
      ? `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`
      : `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="1"/>${ser}<c:axId val="111111111"/><c:axId val="222222222"/></c:barChart><c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx><c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeHTML(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${body}</c:plotArea><c:legend><c:legendPos val="b"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
  }
  function chartAnchorXml(n,a){
    return `<xdr:twoCellAnchor><xdr:from><xdr:col>${a.fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${a.fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${a.toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${a.toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${n+1}" name="Chart ${n}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${n}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
  }
  async function saveWorkbookWithCharts(wb,sheetIndex,charts,filename){
    try{
      if(!window.JSZip||!charts.length)throw new Error("no-jszip");
      const buf=window.XLSX.write(wb,{type:"array",bookType:"xlsx"});
      const zip=await window.JSZip.loadAsync(buf);
      const relEntries=[],anchors=[];
      charts.forEach((c,i)=>{const n=i+1;zip.file(`xl/charts/chart${n}.xml`,chartPartXml(c));relEntries.push(`<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${n}.xml"/>`);anchors.push(chartAnchorXml(n,c.anchor));});
      zip.file("xl/drawings/drawing1.xml",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors.join("")}</xdr:wsDr>`);
      zip.file("xl/drawings/_rels/drawing1.xml.rels",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relEntries.join("")}</Relationships>`);
      const sheetPath=`xl/worksheets/sheet${sheetIndex}.xml`;
      let sheetXml=await zip.file(sheetPath).async("string");
      if(!sheetXml.includes("<drawing "))sheetXml=sheetXml.replace("</worksheet>",`<drawing r:id="rIdD1"/></worksheet>`);
      zip.file(sheetPath,sheetXml);
      const relsPath=`xl/worksheets/_rels/sheet${sheetIndex}.xml.rels`,relsFile=zip.file(relsPath);
      const relEntry=`<Relationship Id="rIdD1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>`;
      let relsXml=relsFile?(await relsFile.async("string")).replace("</Relationships>",`${relEntry}</Relationships>`):`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relEntry}</Relationships>`;
      zip.file(relsPath,relsXml);
      let ctXml=await zip.file("[Content_Types].xml").async("string");
      const overrides=charts.map((c,i)=>`<Override PartName="/xl/charts/chart${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join("")+`<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
      ctXml=ctXml.replace("</Types>",`${overrides}</Types>`);
      zip.file("[Content_Types].xml",ctXml);
      const blob=await zip.generateAsync({type:"blob"});
      const url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),2000);
    }catch(e){console.warn("No fue posible incluir gráficos nativos, se descarga sin gráficos.",e);window.XLSX.writeFile(wb,filename);}
  }
  function donutSVG(percent){
    const p=Math.max(0,Math.min(100,Number(percent)||0)),r=42,c=2*Math.PI*r,off=c*(1-p/100);
    return `<svg viewBox="0 0 100 100"><circle cx="50" cy="50" r="${r}" fill="none" stroke="#e3ddef" stroke-width="14"/><circle cx="50" cy="50" r="${r}" fill="none" stroke="#8064b3" stroke-width="14" stroke-linecap="round" stroke-dasharray="${c.toFixed(2)}" stroke-dashoffset="${off.toFixed(2)}" transform="rotate(-90 50 50)"/></svg>`;
  }
  let selectedAdvisorIds = [];

  // Cache ligero: memoria + localStorage. Los datos se separan por usuario y fecha.
  const CACHE_TTL = 5 * 60 * 1000;
  const memoryCache = new Map();
  const pendingCache = new Map();
  function cacheGet(key){
    const mem=memoryCache.get(key);
    if(mem && Date.now()-mem.ts<CACHE_TTL)return mem.data;
    try{const raw=localStorage.getItem(`tvmax-cache:${key}`);if(!raw)return null;const item=JSON.parse(raw);if(!item||Date.now()-item.ts>=CACHE_TTL){localStorage.removeItem(`tvmax-cache:${key}`);return null;}memoryCache.set(key,item);return item.data;}catch(e){return null;}
  }
  function cacheSet(key,data){const item={ts:Date.now(),data};memoryCache.set(key,item);try{localStorage.setItem(`tvmax-cache:${key}`,JSON.stringify(item));}catch(e){try{localStorage.removeItem(`tvmax-cache:${key}`);}catch(_){}}return data;}
  function cacheInvalidate(prefix){for(const key of memoryCache.keys())if(key.startsWith(prefix))memoryCache.delete(key);try{for(let i=localStorage.length-1;i>=0;i--){const k=localStorage.key(i)||'';if(k.startsWith(`tvmax-cache:${prefix}`))localStorage.removeItem(k);}}catch(e){}}
  async function cachedQuery(key,queryFn){const cached=cacheGet(key);if(cached!==null)return cached;if(pendingCache.has(key))return pendingCache.get(key);const promise=(async()=>{const data=await queryFn();cacheSet(key,data);return data;})().finally(()=>pendingCache.delete(key));pendingCache.set(key,promise);return promise;}
  function monthStartISO(){const n=new Date();return `${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,'0')}-01`;}
  let currentUser = null, currentProfile = null, sales = [], advisors = [], surveys = [], config = { color_principal: "#8b5cf6", logo_url: "" };
  let initializingUserId = null;
  let salesRealtimeChannel = null;
  let adminMonthlySales = [];
  let adminReportSales = null;
  let advisorReportSales = null;
  let advisorReportLoadedKey = "";
  let advisorReportDashboard = null;
  let adminGoalAdvisors = [];
  let surveyReportLoadedKey = "";
  let surveyReportLoading = null;

  document.addEventListener("DOMContentLoaded", async () => {
    bindEvents(); setTodayDefault(); showAuthView(); applyTheme();
    const { data: { session } } = await sbClient.auth.getSession();
    if (session?.user) await initializeSession(session.user);
    sbClient.auth.onAuthStateChange(async (event, session) => {
      if (event === "SIGNED_OUT") { if(salesRealtimeChannel){sbClient.removeChannel(salesRealtimeChannel);salesRealtimeChannel=null;} currentUser = null; currentProfile = null; sales = []; advisors = []; surveys = []; showAuthView(); return; }
      if (session?.user && event !== "INITIAL_SESSION") await initializeSession(session.user);
    });
  });

  function bindEvents() {
    id("login-form").addEventListener("submit", login); id("register-form").addEventListener("submit", registerAdvisor); id("sale-form").addEventListener("submit", registerSale);
    id("btn-show-register").addEventListener("click", () => { id("auth-view").classList.add("hidden"); id("register-view").classList.remove("hidden"); });
    id("btn-back-login").addEventListener("click", showAuthView); id("btn-logout").addEventListener("click", logout);
    id("btn-menu").addEventListener("click", () => id("sidebar").classList.toggle("open")); id("btn-close-menu").addEventListener("click", closeSidebar);
    id("filtroAsesor").addEventListener("input", renderAdvisorTable);
    ["filtroAsesorDesde","filtroAsesorHasta"].forEach(x=>{id(x)?.addEventListener("change",()=>{loadAdvisorSalesForFilters(true).catch(e=>console.warn("No fue posible cargar el periodo del asesor",e));});});
    id("filtro-asesor-usuarios")?.addEventListener("input", debounce(searchAdvisor,350));
    id("btn-refresh-dashboard")?.addEventListener("click", async()=>{cacheInvalidate(`dashboard:tvmax:${currentUser?.id||""}:`);cacheInvalidate(`sales:tvmax:admin:${getTodayISO()}`);cacheInvalidate(`admin-goal-sales:tvmax:${monthStartISO()}`);adminReportSales=null;await loadAdminData();showToast("Dashboard actualizado.");});
    id("btn-refresh-advisor-dashboard")?.addEventListener("click", async()=>{cacheInvalidate(`dashboard:tvmax:${currentUser?.id||""}:`);cacheInvalidate(`sales:tvmax:${currentUser?.id||""}:${getTodayISO()}`);cacheInvalidate(`surveys:tvmax:${currentUser?.id||""}:${getTodayISO()}`);await loadAdvisorData();showToast("Resumen mensual actualizado.");});
    ["filtroAdminTexto","filtroEstadoAdmin","filtroTipoAdmin","filtroServicioAdmin","filtroZonaAdmin"].forEach(x => { id(x).addEventListener("input", renderAdmin); id(x).addEventListener("change", renderAdmin); });
    ["filtroDesdeAdmin","filtroHastaAdmin"].forEach(x => { id(x).addEventListener("change",()=>{loadAdminSalesForFilters(true).catch(e=>{console.warn("No fue posible cargar el periodo de ventas",e);renderAdmin();});}); });
    id("btn-clear-filters").addEventListener("click", clearAdminFilters); id("btn-preview-report").addEventListener("click", async()=>{await loadAdminSalesForFilters(false);previewReport();}); id("btn-close-report-preview").addEventListener("click", closeReportPreview); id("btn-print-report").addEventListener("click", async()=>{await loadAdminSalesForFilters(false);printReport();}); id("btn-pdf-report").addEventListener("click", async()=>{await loadAdminSalesForFilters(false);downloadPDF(buildReportSummaryHTML);}); id("btn-excel-report").addEventListener("click", async()=>{await loadAdminSalesForFilters(false);downloadExcel();});
    id("filtroAsesorAdminBtn").addEventListener("click",(e)=>{e.stopPropagation();id("filtroAsesorAdminPanel").classList.toggle("hidden");});
    id("filtroAsesorAdminAll").addEventListener("click",()=>{selectedAdvisorIds=advisors.map(a=>a.id);syncAdvisorFilterUI();renderAdmin();});
    id("filtroAsesorAdminClear").addEventListener("click",()=>{selectedAdvisorIds=[];syncAdvisorFilterUI();renderAdmin();});
    document.addEventListener("click",(e)=>{const wrap=id("filtroAsesorAdminWrap");if(wrap&&!wrap.contains(e.target))id("filtroAsesorAdminPanel").classList.add("hidden");});
    id("admin-user-form").addEventListener("submit", saveAdminUser); id("btn-cancel-user-edit").addEventListener("click", resetUserForm);
    id("config-form").addEventListener("submit", saveConfig); id("btn-remove-logo").addEventListener("click", removeLogo);
    id("btn-asesor-report").addEventListener("click", previewAdvisorReport); id("btn-asesor-print").addEventListener("click", printAdvisorReport); id("btn-asesor-pdf").addEventListener("click", downloadAdvisorPDF); id("btn-asesor-excel")?.addEventListener("click", downloadAdvisorExcel);
    id("survey-form").addEventListener("submit", registerSurvey);
    ensureSurveyZoneFilter();
    ["filtroEncuestaTexto","filtroEncuestaAsesor","filtroEncuestaQ6","filtroEncuestaZona"].forEach(x=>{id(x)?.addEventListener("input",renderSurveyReport);id(x)?.addEventListener("change",renderSurveyReport);});
    ["filtroEncuestaDesde","filtroEncuestaHasta"].forEach(x=>{id(x)?.addEventListener("change",()=>{loadSurveyReportData(true).catch(e=>console.warn("No fue posible actualizar el rango de encuestas",e));});});
    id("btn-clear-survey-filters").addEventListener("click",clearSurveyFilters);
    id("btn-preview-survey-report").addEventListener("click",()=>previewReport(buildSurveyReportHTML));
    id("btn-print-survey-report").addEventListener("click",()=>printReport(buildSurveyReportHTML));
    id("btn-pdf-survey-report").addEventListener("click",()=>downloadPDF(buildSurveySummaryHTML,"reporte-encuestas"));
    id("btn-excel-survey-report").addEventListener("click",downloadSurveyExcel);
    id("btn-download-backup").addEventListener("click", downloadBackup);
    ensureSurveyZoneFilter();
  }

  async function login(e) { e.preventDefault(); const email=value("login-email"), password=id("login-password").value; setButtonBusy(e.submitter,true,"Ingresando..."); const {data,error}=await sbClient.auth.signInWithPassword({email,password}); setButtonBusy(e.submitter,false,"Ingresar"); if(error){showToast(authError(error),true);return;} await initializeSession(data.user); }

  async function registerAdvisor(e) {
    e.preventDefault(); const password=id("reg-password").value, confirm=id("reg-password-confirm").value;
    if(password!==confirm){showToast("Las contraseñas no coinciden.",true);return;} if(password.length<6){showToast("La contraseña debe tener mínimo 6 caracteres.",true);return;}
    const payload={nombre:value("reg-nombre"),apellido:value("reg-apellido"),documento:value("reg-documento"),telefono:value("reg-telefono"),zona:"",rol:"asesor"};
    setButtonBusy(e.submitter,true,"Registrando..."); const {data,error}=await sbClient.auth.signUp({email:value("reg-email"),password,options:{data:payload}}); setButtonBusy(e.submitter,false,"Registrar asesor");
    if(error){showToast(authError(error),true);return;} id("register-form").reset(); if(data.session){showToast("Asesor registrado correctamente.");await initializeSession(data.user);}else{showToast("Registro creado. Revisa el correo para confirmar la cuenta.");showAuthView();}
  }

  async function initializeSession(user) {
    if(!user?.id)return;
    if(initializingUserId===user.id)return;
    initializingUserId=user.id;
    try{
      currentUser=user;
      const {data:profile,error}=await sbClient.from("perfiles").select("*").eq("id",user.id).single();
      if(error){console.error(error);await sbClient.auth.signOut();showToast("No fue posible cargar tu perfil. Ejecuta el SQL actualizado.",true);return;}
      if(profile.activo === false){await sbClient.auth.signOut();showToast("Tu usuario está inhabilitado. Contacta al administrador.",true);return;}
      currentProfile=profile;
      updateSessionHeader(); buildSidebar();
      // Mostrar la aplicación inmediatamente; las consultas secundarias se ejecutan en segundo plano.
      showView(profile.rol==="administrador"?"admin-dashboard":"vista-asesor");
      subscribeSalesRealtime();
      loadConfig().catch(e=>console.warn("No fue posible cargar la configuración",e));
      if(profile.rol==="administrador")loadAdminData().catch(e=>console.error("Error cargando dashboard de administrador",e));
      else loadAdvisorData().catch(e=>console.error("Error cargando dashboard del asesor",e));
    }finally{initializingUserId=null;}
  }
  async function loadConfig(){const key="config:tvmax";const cached=cacheGet(key);if(cached){config=cached;applyTheme();renderConfig();return;}const {data,error}=await sbClient.from("configuracion").select("color_principal,logo_url").eq("id",1).maybeSingle();if(!error&&data){config=data;cacheSet(key,config);if(currentProfile?.rol==="administrador"&&String(data.logo_url||"").startsWith("data:image/"))migrateLegacyLogo(data.logo_url);}applyTheme();renderConfig();}
  async function migrateLegacyLogo(dataUrl){try{const blob=await fetch(dataUrl).then(r=>r.blob());const ext=(blob.type.split("/")[1]||"png").replace("jpeg","jpg");const path=`tvmax/logo-${Date.now()}.${ext}`;const up=await sbClient.storage.from("app-assets").upload(path,blob,{cacheControl:"31536000",upsert:false,contentType:blob.type});if(up.error)return;const url=sbClient.storage.from("app-assets").getPublicUrl(path).data.publicUrl;const r=await sbClient.from("configuracion").update({logo_url:url,updated_by:currentUser.id}).eq("id",1);if(!r.error){config.logo_url=url;cacheSet("config:tvmax",config);renderConfig();}}catch(e){console.warn("No fue posible migrar el logo anterior al Storage",e);}}
  async function loadMonthlyDashboard(force=false){
    const key=`dashboard:tvmax:v2:${currentUser.id}:${monthStartISO()}`;
    if(force)cacheInvalidate(key);else{const cached=cacheGet(key);if(cached!==null)return cached;}
    const {data,error}=await sbClient.rpc("dashboard_tvmax_mensual",{p_month_start:monthStartISO()});
    if(!error&&data){cacheSet(key,data);return data;}
    console.warn("RPC dashboard_tvmax_mensual no disponible; usando consulta mensual de respaldo",error);
    const inicio=monthStartISO();
    const fin=new Date(); fin.setMonth(fin.getMonth()+1); fin.setDate(0);
    const finISO=`${fin.getFullYear()}-${String(fin.getMonth()+1).padStart(2,'0')}-${String(fin.getDate()).padStart(2,'0')}`;
    const [vr,ar]=await Promise.all([
      sbClient.from("ventas").select("tipo_operacion,estado_instalacion,servicio,asesor_id",{count:"exact"}).gte("fecha_venta",inicio).lte("fecha_venta",finISO),
      sbClient.from("perfiles").select("id,nombre,apellido,email,meta_mensual").eq("rol","asesor").neq("activo",false)
    ]);
    if(vr.error){console.error(vr.error);return null;}
    const rows=vr.data||[], profiles=ar.data||[];
    const services={}; rows.forEach(x=>{if(x.servicio)services[x.servicio]=(services[x.servicio]||0)+1;});
    const result={total:rows.length,ventas:rows.filter(x=>x.tipo_operacion==="Venta").length,reconexiones:rows.filter(x=>x.tipo_operacion==="Reconexión").length,pendientes:rows.filter(x=>x.estado_instalacion==="PENDIENTE").length,realizadas:rows.filter(x=>x.estado_instalacion==="REALIZADA").length,canceladas:rows.filter(x=>x.estado_instalacion==="CANCELADA").length,servicios,asesores:profiles.map(a=>({id:a.id,nombre:a.nombre,apellido:a.apellido,email:a.email,meta:Number(a.meta_mensual)||50,realizadas:rows.filter(x=>x.asesor_id===a.id&&(x.tipo_operacion==="Venta"||x.tipo_operacion==="Reconexión")).length}))};
    cacheSet(key,result);return result;
  }
  async function loadTodaySalesData(){const today=getTodayISO(),uid=currentUser.id;const [sr,qr,dashboard]=await Promise.all([
    cachedQuery(`sales:tvmax:${uid}:${today}`,async()=>{const r=await sbClient.from("ventas").select("id,asesor_id,tipo_operacion,codigo_cliente,codigo_servicio,descripcion_servicio,zona,fecha_venta,estado_instalacion,fecha_instalacion,created_at,updated_at,servicio").eq("asesor_id",uid).eq("fecha_venta",today).order("id",{ascending:false}).limit(10);if(r.error)throw r.error;return r.data||[];}),
    cachedQuery(`surveys:tvmax:${uid}:${today}`,async()=>{const r=await sbClient.from("encuestas").select("id,asesor_id,codigo_nombre_usuario,q2_servicio,observacion_q2,q3_tecnica,observacion_q3,q4_administrativa,observacion_q4,q5_agilidad,observacion_q5,q6_recomendaria,observacion_q6,q7_recomendacion,fecha_encuesta,created_at,updated_at").eq("asesor_id",uid).eq("fecha_encuesta",today).order("id",{ascending:false}).limit(10);if(r.error)throw r.error;return r.data||[];}),loadMonthlyDashboard()]);
    sales=sr;surveys=qr;applyAdvisorProfile();const ma=dashboard?.asesores?.find(a=>a.id===uid);updateAdvisorDashboard(ma);renderAdvisorTable();renderAdvisorSurveys();
  }
  async function loadAdvisorData(){await loadTodaySalesData();}
  async function loadAdminData(){
    const today=getTodayISO();
    // El dashboard mensual se solicita de forma independiente para que pueda pintarse
    // apenas responde el RPC, sin esperar las tablas del día ni otras consultas.
    const dashboardPromise=loadMonthlyDashboard().then(d=>{if(d)updateAdminDashboard(d);return d;}).catch(e=>{console.warn("Dashboard mensual no disponible",e);return null;});
    const results=await Promise.allSettled([
      cachedQuery(`sales:tvmax:admin:${today}`,async()=>{const r=await sbClient.from("ventas").select("id,asesor_id,tipo_operacion,codigo_cliente,codigo_servicio,descripcion_servicio,zona,fecha_venta,estado_instalacion,fecha_instalacion,created_at,updated_at,servicio,perfiles:asesor_id(id,nombre,apellido,zona,email,meta_mensual,activo)").eq("fecha_venta",today).order("id",{ascending:false}).limit(10);if(r.error)throw r.error;return r.data||[];}),
      sbClient.from("perfiles").select("id",{count:"exact",head:true}).eq("rol","asesor"),
      cachedQuery(`surveys:tvmax:admin:${today}`,async()=>{const r=await sbClient.from("encuestas").select("id,asesor_id,codigo_nombre_usuario,q2_servicio,observacion_q2,q3_tecnica,observacion_q3,q4_administrativa,observacion_q4,q5_agilidad,observacion_q5,q6_recomendaria,observacion_q6,q7_recomendacion,fecha_encuesta,created_at,updated_at,perfiles:asesor_id(id,nombre,apellido,email,zona,activo,rol)").eq("fecha_encuesta",today).order("id",{ascending:false}).limit(10);if(r.error)throw r.error;return r.data||[];}),
      dashboardPromise,
      cachedQuery(`admin-goal-sales:tvmax:${monthStartISO()}`,async()=>{const r=await sbClient.from("ventas").select("id,asesor_id,tipo_operacion,fecha_venta").eq("tipo_operacion","Venta").gte("fecha_venta",monthStartISO()).lt("fecha_venta",nextMonthStartISO());if(r.error)throw r.error;return r.data||[];})
    ]);
    const sr=results[0].status==="fulfilled"?results[0].value:[];
    const cnt=results[1].status==="fulfilled"?results[1].value:{count:0};
    const qr=results[2].status==="fulfilled"?results[2].value:[];
    const dashboard=results[3].status==="fulfilled"?results[3].value:null;
    const monthlySales=results[4].status==="fulfilled"?results[4].value:[];
    results.forEach((r,i)=>{if(r.status==="rejected")console.warn("Consulta TV MAX falló",i,r.reason);});
    sales=sr||[];adminMonthlySales=monthlySales||[];advisors=[];surveys=qr||[];
    adminGoalAdvisors=Array.isArray(dashboard?.asesores)?dashboard.asesores:[];
    setText("advisor-count",cnt?.count??0);
    if(dashboard)updateAdminDashboard(dashboard);
    populateAdminFilters();populateSurveyAdvisorFilter();renderAdmin();renderUsers();renderSurveyReport();renderConfig();
    return dashboard;
  }
  async function loadAdvisorsForFilters(){advisors=await cachedQuery("advisors:tvmax:directory",async()=>{const r=await sbClient.from("perfiles").select("id,nombre,apellido,email,zona,meta_mensual,activo,rol").eq("rol","asesor").order("nombre").order("apellido");if(r.error)throw r.error;return r.data||[];});populateAdminFilters();populateSurveyAdvisorFilter();}
  async function searchAdvisor(){const q=value("filtro-asesor-usuarios");if(!q){advisors=[];renderUsers();return;}const safe=q.replace(/[(),*]/g," ").replace(/\s+/g," ").trim();const r=await sbClient.from("perfiles").select("id,nombre,apellido,documento,telefono,zona,email,rol,meta_mensual,activo").eq("rol","asesor").or(`nombre.ilike.%${safe}%,apellido.ilike.%${safe}%,email.ilike.%${safe}%,documento.ilike.%${safe}%`).order("nombre").limit(1);if(r.error){showToast("No fue posible buscar el asesor.",true);return;}advisors=r.data||[];renderUsers();}

  async function registerSale(e){
    e.preventDefault(); if(!currentUser||!currentProfile){showToast("Tu sesión no está disponible.",true);return;}
    const row={asesor_id:currentUser.id,tipo_operacion:value("tipoOperacion"),codigo_cliente:value("codigoCliente"),servicio:value("servicio"),descripcion_servicio:value("descripcionServicio"),zona:value("zona"),fecha_venta:id("fechaVenta").value,estado_instalacion:"PENDIENTE",fecha_instalacion:null};
    if(!row.tipo_operacion||!row.codigo_cliente||!row.servicio||!row.descripcion_servicio||!row.zona||!row.fecha_venta){showToast("Completa todos los campos obligatorios.",true);return;}
    const {data,error}=await sbClient.from("ventas").insert(row).select().single(); if(error){console.error(error);showToast(error.message||"No fue posible registrar la operación.",true);return;}
    e.target.reset();applyAdvisorProfile();setTodayDefault();sales=[data,...sales].slice(0,10);cacheInvalidate(`sales:tvmax:${currentUser.id}:${getTodayISO()}`);cacheInvalidate(`dashboard:tvmax:${currentUser.id}:${monthStartISO()}`);renderAdvisorTable();loadMonthlyDashboard(true).then(d=>updateAdvisorDashboard(d?.asesores?.find(a=>a.id===currentUser.id)));showToast(`${row.tipo_operacion} registrada correctamente.`);
  }


  async function registerSurvey(e){
    e.preventDefault();
    if(!currentUser||!["asesor","administrador"].includes(currentProfile?.rol)){showToast("No tienes permisos para registrar encuestas.",true);return;}
    const row={
      asesor_id:currentUser.id,
      codigo_nombre_usuario:value("enc-codigo-nombre"),
      q2_servicio:document.querySelector('input[name="enc-q2"]:checked')?.value||"",
      observacion_q2:value("enc-obs2"),
      q3_tecnica:document.querySelector('input[name="enc-q3"]:checked')?.value||"",
      observacion_q3:value("enc-obs3"),
      q4_administrativa:document.querySelector('input[name="enc-q4"]:checked')?.value||"",
      observacion_q4:value("enc-obs4"),
      q5_agilidad:document.querySelector('input[name="enc-q5"]:checked')?.value||"",
      observacion_q5:value("enc-obs5"),
      q6_recomendaria:document.querySelector('input[name="enc-q6"]:checked')?.value||"",
      observacion_q6:value("enc-obs6"),
      q7_recomendacion:value("enc-q7"),
      fecha_encuesta:getTodayISO()
    };
    if(!row.codigo_nombre_usuario||!row.q2_servicio||!row.q3_tecnica||!row.q4_administrativa||!row.q5_agilidad||!row.q6_recomendaria){
      showToast("Completa las preguntas obligatorias de la encuesta.",true);return;
    }
    const {data,error}=await sbClient.from("encuestas").insert(row).select("*").single();
    if(error){console.error(error);showToast(error.message||"No fue posible guardar la encuesta.",true);return;}
    e.target.reset();
    surveys.unshift(data);
    renderAdvisorSurveys();
    if(currentProfile?.rol==="administrador") renderSurveyReport();
    showToast("Encuesta guardada correctamente.");
  }

  function renderAdvisorSurveys(){
    const tabla=id("tabla-encuestas-asesor"); if(!tabla)return;
    const mine=surveys.filter(s=>s.asesor_id===currentUser?.id);
    setText("advisor-survey-count",`${mine.length} encuesta${mine.length===1?"":"s"}`);
    tabla.innerHTML=mine.length?mine.map(s=>`<tr>
      <td>${formatDate(s.fecha_encuesta)}</td>
      <td>${escapeHTML(s.codigo_nombre_usuario||"—")}</td>
      <td>${escapeHTML(s.q2_servicio||"—")}</td>
      <td>${escapeHTML(s.q3_tecnica||"—")}</td>
      <td>${escapeHTML(s.q4_administrativa||"—")}</td>
      <td>${escapeHTML(s.q5_agilidad||"—")}</td>
      <td>${escapeHTML(s.q6_recomendaria||"—")}</td>
      <td>${escapeHTML(s.q7_recomendacion||"—")}</td>
    </tr>`).join(""):`<tr class="empty-row"><td colspan="10">Aún no has registrado encuestas.</td></tr>`;
  }

  function subscribeSalesRealtime(){
    if(!currentUser)return;
    if(salesRealtimeChannel)sbClient.removeChannel(salesRealtimeChannel);
    salesRealtimeChannel=sbClient.channel(`tvmax-ventas-${currentUser.id}`)
      .on("postgres_changes",{event:"*",schema:"public",table:"ventas"},async(payload)=>{
        cacheInvalidate(`dashboard:tvmax:${currentUser.id}:`);
        cacheInvalidate(`dashboard:tvmax:v2:${currentUser.id}:`);
        const affected=payload.new||payload.old||{};
        const mine=currentProfile?.rol==="asesor"&&String(affected.asesor_id||"")===String(currentUser.id);
        const today=getTodayISO();
        if(mine && String(affected.fecha_venta||"")===today){
          cacheInvalidate(`sales:tvmax:${currentUser.id}:${today}`);
          const r=await sbClient.from("ventas").select("id,asesor_id,tipo_operacion,codigo_cliente,codigo_servicio,descripcion_servicio,zona,fecha_venta,estado_instalacion,fecha_instalacion,created_at,updated_at,servicio").eq("asesor_id",currentUser.id).eq("fecha_venta",today).order("id",{ascending:false}).limit(10);
          if(!r.error)sales=r.data||[];
        }
        const dashboard=await loadMonthlyDashboard(true);
        if(currentProfile?.rol==="asesor"){const ma=dashboard?.asesores?.find(a=>a.id===currentUser.id);updateAdvisorDashboard(ma);renderAdvisorTable();}
        else if(dashboard){adminGoalAdvisors=Array.isArray(dashboard.asesores)?dashboard.asesores:[];updateAdminDashboard(dashboard);renderAdmin();}
      })
      .subscribe();
  }

  async function setInstallation(id,state){
    if(!currentProfile||currentProfile.rol!=="administrador")return;
    if(state==="CANCELADA"){
      const {data,error}=await sbClient.from("ventas").update({estado_instalacion:"CANCELADA",fecha_instalacion:null}).eq("id",id).select(`*,perfiles:asesor_id (id,nombre,apellido,zona,email,meta_mensual,activo)`).single(); if(error){showToast("No fue posible cancelar la operación.",true);return;} await updateSaleLocal(data);showToast("Operación marcada como cancelada.");return;
    }
    const input=document.getElementById(`date-${id}`); if(!input?.value){showToast("Selecciona la fecha de instalación.",true);return;}
    const {data,error}=await sbClient.from("ventas").update({fecha_instalacion:input.value,estado_instalacion:"REALIZADA"}).eq("id",id).select(`*,perfiles:asesor_id (id,nombre,apellido,zona,email,meta_mensual,activo)`).single(); if(error){showToast("No fue posible actualizar la instalación.",true);return;} await updateSaleLocal(data);showToast("Instalación marcada como realizada.");
  }
  async function updateSaleLocal(data){
    const i=sales.findIndex(x=>x.id===data.id);if(i>=0)sales[i]=data;
    renderAdmin();
    cacheInvalidate(`dashboard:tvmax:${currentUser?.id||""}:`);
    cacheInvalidate(`dashboard:tvmax:v2:${currentUser?.id||""}:`);
    const dashboard=await loadMonthlyDashboard(true);
    if(dashboard)updateAdminDashboard(dashboard);
  }

  async function deleteSale(id){if(!confirm("¿Eliminar definitivamente esta operación? Esta acción no se puede deshacer."))return;const {error}=await sbClient.from("ventas").delete().eq("id",id);if(error){showToast("No fue posible eliminar la venta. Verifica las políticas RLS.",true);return;}sales=sales.filter(x=>x.id!==id);renderAdmin();updateAdminDashboard();showToast("Operación eliminada.");}

  function buildSidebar(){const nav=id("sidebar-nav");const admin=currentProfile?.rol==="administrador";const items=admin?[ ["admin-dashboard","▦","Dashboard"],["vista-admin","▤","Operaciones"],["vista-encuestas","☑","Encuesta"],["vista-reporte-encuestas","▤","Reporte de encuestas"],["vista-usuarios","♙","Usuarios"],["vista-configuracion","⚙","Configuración"],["vista-respaldo","⭳","Respaldo"] ]:[["vista-asesor","▦","Mi dashboard"],["vista-asesor","＋","Registrar operación"],["vista-asesor","▤","Mis operaciones"],["vista-encuestas","☑","Encuestas"]];nav.innerHTML=items.map(([target,icon,label])=>`<button class="nav-item" type="button" data-target="${target}" data-anchor="${target==='vista-asesor'?label:''}"><span>${icon}</span>${label}</button>`).join("");nav.querySelectorAll(".nav-item").forEach(b=>b.addEventListener("click",()=>{showView(b.dataset.target);if(b.dataset.anchor==="Registrar operación")id("asesor-form-section").scrollIntoView({behavior:"smooth"});if(b.dataset.anchor==="Mis operaciones")document.querySelector("#vista-asesor .table-card").scrollIntoView({behavior:"smooth"});if(b.dataset.target==="vista-encuestas")renderAdvisorSurveys();closeSidebar();}));}
  function closeSidebar(){id("sidebar").classList.remove("open");}

  function updateSessionHeader(){const name=[currentProfile?.nombre,currentProfile?.apellido].filter(Boolean).join(" ")||"Usuario", role=currentProfile?.rol==="administrador"?"Administrador":"Asesor";id("user-name").textContent=name;id("user-role").textContent=role;id("user-avatar").textContent=name.charAt(0).toUpperCase();id("sidebar-user-name").textContent=name;id("sidebar-user-role").textContent=role;id("session-area").classList.remove("hidden");id("btn-menu").classList.remove("hidden");id("sidebar").classList.remove("hidden");if(id("survey-mode-description"))id("survey-mode-description").textContent=currentProfile?.rol==="administrador"?"Diligencia una encuesta de satisfacción como mecanismo de control y seguimiento de la atención al usuario.":"Diligencia la encuesta utilizando exactamente las preguntas del formulario de satisfacción de Grupo TV Max.";}
  function applyAdvisorProfile(){const zona=currentProfile?.zona||"";id("zona").value=zona;id("asesor-zone-badge").textContent=`Zona: ${zona||"Sin asignar"}`;id("asesor-welcome").textContent=`Registra operaciones y consulta tu avance. Zona asignada: ${zona||"sin asignar"}.`;}

  function getFilteredAdvisorSales(){const source=advisorReportSales||sales,filtro=value("filtroAsesor").toLowerCase(),from=value("filtroAsesorDesde"),to=value("filtroAsesorHasta");return source.filter(s=>{const search=[s.tipo_operacion,s.codigo_cliente,s.servicio,s.descripcion_servicio,s.zona,s.fecha_venta,s.estado_instalacion].join(" ").toLowerCase();return(!filtro||search.includes(filtro))&&(!from||s.fecha_venta>=from)&&(!to||s.fecha_venta<=to);});}
  function renderAdvisorTable(){const tabla=id("tabla-asesor"),filtered=getFilteredAdvisorSales();tabla.innerHTML=filtered.length?filtered.map(s=>`<tr><td>#${s.id}</td><td>${operationBadge(s.tipo_operacion)}</td><td>${escapeHTML(s.codigo_cliente)}</td><td>${serviceBadge(s.servicio)}</td><td>${escapeHTML(s.descripcion_servicio)}</td><td>${escapeHTML(s.zona)}</td><td>${formatDate(s.fecha_venta)}</td><td>${installationStatus(s.estado_instalacion)}</td><td>${formatDate(s.fecha_instalacion)}</td></tr>`).join(""):`<tr class="empty-row"><td colspan="9">${sales.length?"No se encontraron operaciones.":"No hay operaciones registradas."}</td></tr>`;}
  function updateAdvisorStats(monthlyAdvisor=null){
    const m=monthlyAdvisor||{};
    const total=Number(m.total)||0,ventas=Number(m.ventas)||0,recon=Number(m.reconexiones)||0,real=Number(m.realizadas_estado)||0;
    setText("asesor-total-count",total);setText("asesor-ventas-count",ventas);setText("asesor-reconexion-count",recon);setText("asesor-complete-count",real);
  }
  function updateAdvisorDashboard(monthlyAdvisor=null){
    updateAdvisorStats(monthlyAdvisor);
    const meta=Math.max(1,Number(monthlyAdvisor?.meta)||Number(currentProfile?.meta_mensual)||50),made=Number(monthlyAdvisor?.realizadas)||0,pct=Math.min(100,Math.round(made/meta*100));
    setText("asesor-goal-title",`${made} / ${meta} operaciones`);
    setText("asesor-goal-detail",`Avance mensual · ${new Date().toLocaleDateString("es-CO",{month:"long",year:"numeric"})}. Las tarjetas muestran el acumulado del mes.`);
    if(id("asesor-goal-bar"))id("asesor-goal-bar").style.width=`${pct}%`;setText("asesor-goal-percent",`${pct}%`);
  }

  function renderAdmin(){const tabla=id("tabla-admin");if(!tabla)return;const filtered=getFilteredAdminSales();tabla.innerHTML=filtered.length?filtered.map(s=>{const a=s.perfiles||{};const name=[a.nombre,a.apellido].filter(Boolean).join(" ")||"—";return `<tr><td>#${s.id}</td><td>${escapeHTML(name)}</td><td>${operationBadge(s.tipo_operacion)}</td><td>${escapeHTML(s.codigo_cliente)}</td><td>${serviceBadge(s.servicio)}</td><td>${escapeHTML(s.descripcion_servicio)}</td><td>${escapeHTML(s.zona)}</td><td>${formatDate(s.fecha_venta)}</td><td>${installationStatus(s.estado_instalacion)}</td><td><input class="installation-date" type="date" id="date-${s.id}" value="${s.fecha_instalacion||""}" ${s.estado_instalacion!=="PENDIENTE"?"disabled":""}></td><td class="action-cell"><button class="btn-save-installation" ${s.estado_instalacion!=="PENDIENTE"?"disabled":""} onclick="setInstallation(${s.id},'REALIZADA')">Realizar</button><button class="btn-cancel-sale" ${s.estado_instalacion!=="PENDIENTE"?"disabled":""} onclick="setInstallation(${s.id},'CANCELADA')">Cancelar</button><button class="btn-delete" onclick="deleteSale(${s.id})">Eliminar</button></td></tr>`;}).join(""):`<tr class="empty-row"><td colspan="11">${sales.length?"No se encontraron operaciones con los filtros seleccionados.":"No hay operaciones registradas."}</td></tr>`;id("admin-result-count").textContent=`${filtered.length} resultado${filtered.length===1?"":"s"}`;}
  function getFilteredAdminSales(){const text=value("filtroAdminTexto").toLowerCase(),state=id("filtroEstadoAdmin").value,type=id("filtroTipoAdmin").value,service=id("filtroServicioAdmin").value,zone=id("filtroZonaAdmin").value,from=id("filtroDesdeAdmin").value,to=id("filtroHastaAdmin").value;return (adminReportSales||sales).filter(s=>{const a=s.perfiles||{},search=[a.nombre,a.apellido,a.email,s.tipo_operacion,s.codigo_cliente,s.servicio,s.descripcion_servicio,s.zona,s.fecha_venta].join(" ").toLowerCase();return(!text||search.includes(text))&&(!selectedAdvisorIds.length||selectedAdvisorIds.includes(s.asesor_id))&&(!state||s.estado_instalacion===state)&&(!type||s.tipo_operacion===type)&&(!service||s.servicio===service)&&(!zone||s.zona===zone)&&(!from||s.fecha_venta>=from)&&(!to||s.fecha_venta<=to);});}
  function populateAdminFilters(){
    selectedAdvisorIds=selectedAdvisorIds.filter(id=>advisors.some(a=>a.id===id));
    const opts=id("filtroAsesorAdminOptions");
    opts.innerHTML=advisors.map(a=>{const n=escapeHTML([a.nombre,a.apellido].filter(Boolean).join(" ")||a.email);return `<label class="multi-select-option"><input type="checkbox" value="${a.id}" ${selectedAdvisorIds.includes(a.id)?"checked":""}><span>${n}</span></label>`;}).join("")||'<p class="muted">No hay asesores registrados.</p>';
    opts.querySelectorAll('input[type="checkbox"]').forEach(cb=>cb.addEventListener("change",()=>{
      selectedAdvisorIds=[...opts.querySelectorAll('input[type="checkbox"]:checked')].map(x=>x.value);
      syncAdvisorFilterUI();renderAdmin();
    }));
    syncAdvisorFilterUI();
    const zone=id("filtroZonaAdmin"),zVal=zone.value;const zones=[...new Set(sales.map(s=>s.zona).filter(Boolean))].sort((a,b)=>a.localeCompare(b));zone.innerHTML='<option value="">Todas las zonas</option>'+zones.map(z=>`<option>${escapeHTML(z)}</option>`).join("");zone.value=zVal;
  }
  function syncAdvisorFilterUI(){
    const btn=id("filtroAsesorAdminBtn"); if(!btn)return;
    if(!selectedAdvisorIds.length){btn.textContent="Todos los asesores";}
    else if(selectedAdvisorIds.length===1){const a=advisors.find(x=>x.id===selectedAdvisorIds[0]);btn.textContent=a?([a.nombre,a.apellido].filter(Boolean).join(" ")||a.email):"1 asesor seleccionado";}
    else{btn.textContent=`${selectedAdvisorIds.length} asesores seleccionados`;}
    const opts=id("filtroAsesorAdminOptions"); if(opts)opts.querySelectorAll('input[type="checkbox"]').forEach(cb=>{cb.checked=selectedAdvisorIds.includes(cb.value);});
  }

  function surveyReportPeople(){
    const map=new Map(advisors.map(a=>[a.id,a]));
    surveys.forEach(s=>{if(s.perfiles?.id&&!map.has(s.perfiles.id))map.set(s.perfiles.id,s.perfiles);});
    const people=[...map.values()];
    if(currentProfile?.rol==="administrador" && currentProfile?.id && !people.some(p=>p.id===currentProfile.id)){people.push({...currentProfile});}
    return people;
  }

  function populateSurveyAdvisorFilter(){
    const el=id("filtroEncuestaAsesor"); if(!el)return;
    const selected=el.value;
    el.innerHTML='<option value="">Todos</option>'+surveyReportPeople().map(a=>`<option value="${a.id}">${escapeHTML([a.nombre,a.apellido].filter(Boolean).join(" ")||a.email||(a.rol==="administrador"?"Administrador":"Asesor"))}${a.rol==="administrador"?" · Administrador":""}</option>`).join("");
    el.value=selected;
  }

  function getFilteredSurveys(){
    const text=value("filtroEncuestaTexto").toLowerCase();
    const advisor=value("filtroEncuestaAsesor");
    const recommend=value("filtroEncuestaQ6");
    const zone=value("filtroEncuestaZona");
    const from=value("filtroEncuestaDesde");
    const to=value("filtroEncuestaHasta");
    return surveys.filter(s=>{
      const a=s.perfiles||{};
      const name=[a.nombre,a.apellido].filter(Boolean).join(" ");
      const matchesText=!text||[s.codigo_nombre_usuario,s.q2_servicio,s.q3_tecnica,s.q4_administrativa,s.q5_agilidad,s.q6_recomendaria,s.q7_recomendacion,name].join(" ").toLowerCase().includes(text);
      const matchesAdvisor=!advisor||s.asesor_id===advisor;
      const matchesRecommend=!recommend||s.q6_recomendaria===recommend;
      const matchesZone=!zone||String(s.perfiles?.zona||"")===zone;
      const matchesFrom=!from||String(s.fecha_encuesta||"")>=from;
      const matchesTo=!to||String(s.fecha_encuesta||"")<=to;
      return matchesText&&matchesAdvisor&&matchesRecommend&&matchesZone&&matchesFrom&&matchesTo;
    });
  }

  function renderSurveyReport(){
    const list=getFilteredSurveys(),tabla=id("tabla-reporte-encuestas"); if(!tabla)return;
    setText("survey-result-count",`${list.length} resultado${list.length===1?"":"s"}`);
    setText("survey-total-count",list.length);
    const yes=list.filter(s=>s.q6_recomendaria==="SI").length;
    setText("survey-recommend-percent",`${list.length?Math.round(yes/list.length*100):0}%`);
    const advisorIds=new Set(list.map(s=>s.asesor_id).filter(Boolean));
    setText("survey-advisor-count",advisorIds.size);
    const reportPeople=surveyReportPeople();
    id("survey-advisor-chart").innerHTML=reportPeople.length?reportPeople.map(a=>{
      const rows=list.filter(s=>s.asesor_id===a.id), yesA=rows.filter(s=>s.q6_recomendaria==="SI").length;
      const pct=rows.length?Math.round(yesA/rows.length*100):0;
      const name=[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email||"Asesor";
      return `<div class="survey-advisor-row"><div class="survey-advisor-head"><strong>${escapeHTML(name)}</strong><span>${rows.length} encuesta${rows.length===1?"":"s"} · ${pct}% recomienda</span></div><div class="survey-advisor-track"><span style="width:${pct}%"></span></div></div>`;
    }).join(""):'<p class="muted">No hay asesores registrados.</p>';
    const zoneCounts={};list.forEach(s=>{const z=s.perfiles?.zona||"Sin zona";zoneCounts[z]=(zoneCounts[z]||0)+1;});
    const zoneSummary=Object.entries(zoneCounts).sort((a,b)=>b[1]-a[1]).map(([z,n])=>`<div class="survey-advisor-row"><div class="survey-advisor-head"><strong>${escapeHTML(z)}</strong><span>${n} encuesta${n===1?"":"s"}</span></div><div class="survey-advisor-track"><span style="width:${total?Math.round(n/total*100):0}%"></span></div></div>`).join("")||'<p class="muted">No hay datos por zona.</p>';
    id("survey-advisor-chart").insertAdjacentHTML("beforeend",`<div class="survey-zone-summary"><span class="section-kicker">POR ZONA</span><h3>Encuestas realizadas por zona</h3>${zoneSummary}</div>`);
    tabla.innerHTML=list.length?list.map(s=>{
      const a=s.perfiles||{},name=[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email||"—";
      return `<tr><td>${formatDate(s.fecha_encuesta)}</td><td>${escapeHTML(name)}</td><td>${escapeHTML(s.perfiles?.zona||"—")}</td><td>${escapeHTML(s.codigo_nombre_usuario||"—")}</td><td>${escapeHTML(s.q2_servicio||"—")}</td><td>${escapeHTML(s.observacion_q2||"—")}</td><td>${escapeHTML(s.q3_tecnica||"—")}</td><td>${escapeHTML(s.observacion_q3||"—")}</td><td>${escapeHTML(s.q4_administrativa||"—")}</td><td>${escapeHTML(s.observacion_q4||"—")}</td><td>${escapeHTML(s.q5_agilidad||"—")}</td><td>${escapeHTML(s.q6_recomendaria||"—")}</td><td>${escapeHTML(s.q7_recomendacion||"—")}</td></tr>`;
    }).join(""):`<tr class="empty-row"><td colspan="13">${surveys.length?"No se encontraron encuestas con los filtros seleccionados.":"No hay encuestas registradas."}</td></tr>`;
  }

  function clearSurveyFilters(){
    ["filtroEncuestaTexto","filtroEncuestaAsesor","filtroEncuestaQ6","filtroEncuestaZona","filtroEncuestaDesde","filtroEncuestaHasta"].forEach(x=>{if(id(x))id(x).value="";});
    loadSurveyReportData(true).catch(e=>console.warn("No fue posible restablecer el reporte de encuestas",e));
  }

  function buildSurveyReportHTML(){
    const list=getFilteredSurveys(),total=list.length,yes=list.filter(s=>s.q6_recomendaria==="SI").length,no=list.filter(s=>s.q6_recomendaria==="NO").length;
    const pct=n=>total?Math.round(n/total*100):0;
    const advisorRows=surveyReportPeople().map(a=>{
      const rows=list.filter(s=>s.asesor_id===a.id), y=rows.filter(s=>s.q6_recomendaria==="SI").length, p=rows.length?Math.round(y/rows.length*100):0;
      if(!rows.length)return "";
      const name=[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email||"Asesor";
      return `<tr><td>${escapeHTML(name)}</td><td>${rows.length}</td><td>${y}</td><td>${rows.length-y}</td><td>${p}%</td></tr>`;
    }).join("")||'<tr><td colspan="5" class="print-empty-row">No hay datos.</td></tr>';
    const detail=list.map(s=>{
      const a=s.perfiles||{},name=[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email||"—";
      const withNote=(val,note)=>escapeHTML(val||"—")+(note?`<br><small>${escapeHTML(note)}</small>`:"");
      return `<tr><td>${formatDate(s.fecha_encuesta)}</td><td>${escapeHTML(name)}</td><td>${escapeHTML(s.perfiles?.zona||"—")}</td><td>${escapeHTML(s.codigo_nombre_usuario||"—")}</td><td>${withNote(s.q2_servicio,s.observacion_q2)}</td><td>${withNote(s.q3_tecnica,s.observacion_q3)}</td><td>${withNote(s.q4_administrativa,s.observacion_q4)}</td><td>${escapeHTML(s.q5_agilidad||"—")}</td><td>${escapeHTML(s.q6_recomendaria||"—")}</td><td>${escapeHTML(s.q7_recomendacion||"—")}</td></tr>`;
    }).join("")||'<tr><td colspan="9" class="print-empty-row">No hay encuestas para mostrar.</td></tr>';
    const dist=(field,opts)=>opts.map(o=>`<tr><td>${escapeHTML(o)}</td><td>${list.filter(s=>s[field]===o).length}</td><td>${pct(list.filter(s=>s[field]===o).length)}%</td></tr>`).join("");
    const period=value("filtroEncuestaDesde")||value("filtroEncuestaHasta")?`${value("filtroEncuestaDesde")?formatDate(value("filtroEncuestaDesde")):"Inicio"} – ${value("filtroEncuestaHasta")?formatDate(value("filtroEncuestaHasta")):"Actual"}`:"Todos los periodos";
    return `<div class="print-report-sheet survey-print-sheet">${config.logo_url?`<div class="print-logo"><img loading="lazy" src="${config.logo_url}" alt="Logo"></div>`:""}<div class="print-header"><div><span class="print-kicker">REPORTE DE ENCUESTAS</span><h1>Satisfacción de usuarios</h1><p>Periodo: <strong>${escapeHTML(period)}</strong></p></div><div class="print-generated">Generado: ${new Date().toLocaleString("es-CO")}</div></div><div class="print-summary"><div class="print-summary-card"><span>Total encuestas</span><strong>${total}</strong></div><div class="print-summary-card"><span>Recomiendan</span><strong>${yes} (${pct(yes)}%)</strong></div><div class="print-summary-card"><span>No recomiendan</span><strong>${no} (${pct(no)}%)</strong></div></div><section class="print-table-section"><div class="print-table-title"><div><span class="print-kicker">POR ASESOR</span><h2>Encuestas registradas por asesor</h2></div></div><div class="print-table-scroll"><table><thead><tr><th>Asesor</th><th>Total</th><th>Sí</th><th>No</th><th>% Sí</th></tr></thead><tbody>${advisorRows}</tbody></table></div></section><section class="print-table-section"><div class="print-table-title"><div><span class="print-kicker">POR ZONA</span><h2>Encuestas realizadas por zona</h2></div></div><div class="print-table-scroll"><table><thead><tr><th>Zona</th><th>Encuestas</th><th>%</th></tr></thead><tbody>${Object.entries(list.reduce((m,s)=>{const z=s.perfiles?.zona||"Sin zona";m[z]=(m[z]||0)+1;return m;},{})).sort((a,b)=>b[1]-a[1]).map(([z,n])=>`<tr><td>${escapeHTML(z)}</td><td>${n}</td><td>${pct(n)}%</td></tr>`).join("")||'<tr><td colspan="3" class="print-empty-row">No hay datos por zona.</td></tr>'}</tbody></table></div></section><section class="print-table-section"><div class="print-table-title"><div><span class="print-kicker">DISTRIBUCIÓN</span><h2>Respuestas por pregunta</h2></div></div><div class="survey-print-distributions"><div><h3>Pregunta 2</h3><p class="q-text">${escapeHTML(SURVEY_QUESTIONS.q2_servicio)}</p><table><thead><tr><th>Respuesta</th><th>Cantidad</th><th>%</th></tr></thead><tbody>${dist("q2_servicio",["BUENO","EXCELENTE","MALO","MUY MALO","REGULAR"])}</tbody></table></div><div><h3>Pregunta 3</h3><p class="q-text">${escapeHTML(SURVEY_QUESTIONS.q3_tecnica)}</p><table><thead><tr><th>Respuesta</th><th>Cantidad</th><th>%</th></tr></thead><tbody>${dist("q3_tecnica",["BUENO","EXCELENTE","MALO","MUY MALO","REGULAR"])}</tbody></table></div><div><h3>Pregunta 4</h3><p class="q-text">${escapeHTML(SURVEY_QUESTIONS.q4_administrativa)}</p><table><thead><tr><th>Respuesta</th><th>Cantidad</th><th>%</th></tr></thead><tbody>${dist("q4_administrativa",["BUENO","EXCELENTE","MALO","MUY MALO","REGULAR"])}</tbody></table></div><div><h3>Pregunta 5</h3><p class="q-text">${escapeHTML(SURVEY_QUESTIONS.q5_agilidad)}</p><table><thead><tr><th>Respuesta</th><th>Cantidad</th><th>%</th></tr></thead><tbody>${dist("q5_agilidad",["AGIL","DEMORADOS","MUY DEMORADOS","NI DEMORADOS NI AGIL"])}</tbody></table></div><div><h3>Pregunta 6</h3><p class="q-text">${escapeHTML(SURVEY_QUESTIONS.q6_recomendaria)}</p><table><thead><tr><th>Respuesta</th><th>Cantidad</th><th>%</th></tr></thead><tbody>${dist("q6_recomendaria",["SI","NO"])}</tbody></table></div></div></section><section class="print-table-section"><div class="print-table-title"><div><span class="print-kicker">DETALLE</span><h2>Respuestas de las encuestas</h2></div><strong>${total} resultado${total===1?"":"s"}</strong></div><div class="print-table-scroll"><table><thead><tr><th>Fecha</th><th>Asesor</th><th>Zona</th><th>Usuario</th><th>Q2 Servicio</th><th>Q3 Técnica</th><th>Q4 Administrativa</th><th>Q5 Agilidad</th><th>Q6 Recomienda</th><th>Q7 Recomendación / felicitación</th></tr></thead><tbody>${detail}</tbody></table></div></section></div>`;
  }

  async function downloadSurveyExcel(){
    if(!window.XLSX){showToast("No se pudo cargar el módulo de Excel.",true);return;}
    const list=getFilteredSurveys();
    try{
      const total=list.length,yes=list.filter(s=>s.q6_recomendaria==="SI").length,no=list.filter(s=>s.q6_recomendaria==="NO").length;
      const pct=n=>total?Math.round(n/total*100):0;
      const period=value("filtroEncuestaDesde")||value("filtroEncuestaHasta")?`${value("filtroEncuestaDesde")?formatDate(value("filtroEncuestaDesde")):"Inicio"} – ${value("filtroEncuestaHasta")?formatDate(value("filtroEncuestaHasta")):"Actual"}`:"Todos los periodos";
      const distCounts=(field,opts)=>opts.map(o=>list.filter(s=>s[field]===o).length);
      const distRows=(field,opts)=>opts.map((o,i)=>[o,distCounts(field,opts)[i],`${pct(distCounts(field,opts)[i])}%`]);
      const advisorRows=surveyReportPeople().map(a=>{const rows=list.filter(s=>s.asesor_id===a.id),y=rows.filter(s=>s.q6_recomendaria==="SI").length;return [[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email||"Asesor",rows.length,y,rows.length-y,`${rows.length?Math.round(y/rows.length*100):0}%`];}).filter(r=>r[1]>0);
      const sheetName="Resumen del informe";
      const resumen=[];const push=r=>resumen.push(r);
      push(["REPORTE DE ENCUESTAS DE SATISFACCIÓN"]);push(["Periodo",period]);push([]);
      push(["RESUMEN GENERAL"]);push(["Indicador","Cantidad","Porcentaje"]);push(["Total encuestas",total,"100%"]);push(["Recomiendan",yes,`${pct(yes)}%`]);push(["No recomiendan",no,`${pct(no)}%`]);push([]);
      push(["POR ASESOR"]);push(["Asesor","Total","Sí","No","% Sí"]);
      (advisorRows.length?advisorRows:[["Sin datos","","","",""]]).forEach(push);push([]);
      push(["POR ZONA"]);push(["Zona","Encuestas","%"]);
      const zoneRows=Object.entries(list.reduce((m,s)=>{const z=s.perfiles?.zona||"Sin zona";m[z]=(m[z]||0)+1;return m;},{})).sort((a,b)=>b[1]-a[1]).map(([z,n])=>[z,n,`${pct(n)}%`]);
      (zoneRows.length?zoneRows:[["Sin datos","",""]]).forEach(push);push([]);
      push(["DISTRIBUCIÓN DE RESPUESTAS POR PREGUNTA"]);push([]);
      push(["PREGUNTA 2",SURVEY_QUESTIONS.q2_servicio]);push(["Respuesta","Cantidad","%"]);
      const q2Opts=["BUENO","EXCELENTE","MALO","MUY MALO","REGULAR"],q2Row=resumen.length;distRows("q2_servicio",q2Opts).forEach(push);push([]);
      push(["PREGUNTA 3",SURVEY_QUESTIONS.q3_tecnica]);push(["Respuesta","Cantidad","%"]);distRows("q3_tecnica",["BUENO","EXCELENTE","MALO","MUY MALO","REGULAR"]).forEach(push);push([]);
      push(["PREGUNTA 4",SURVEY_QUESTIONS.q4_administrativa]);push(["Respuesta","Cantidad","%"]);distRows("q4_administrativa",["BUENO","EXCELENTE","MALO","MUY MALO","REGULAR"]).forEach(push);push([]);
      push(["PREGUNTA 5",SURVEY_QUESTIONS.q5_agilidad]);push(["Respuesta","Cantidad","%"]);distRows("q5_agilidad",["AGIL","DEMORADOS","MUY DEMORADOS","NI DEMORADOS NI AGIL"]).forEach(push);push([]);
      push(["PREGUNTA 6",SURVEY_QUESTIONS.q6_recomendaria]);push(["Respuesta","Cantidad","%"]);
      const q6Row=resumen.length;distRows("q6_recomendaria",["SI","NO"]).forEach(push);push([]);

      const wr=window.XLSX.utils.aoa_to_sheet(resumen);wr["!cols"]=[{wch:34},{wch:70},{wch:15},{wch:15}];
      const wb=window.XLSX.utils.book_new();window.XLSX.utils.book_append_sheet(wb,wr,sheetName);
      const detail=list.map(s=>{const a=s.perfiles||{};return {
        "Fecha":s.fecha_encuesta||"","Asesor":[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email||"","Zona":a.zona||"","Código y nombre del usuario":s.codigo_nombre_usuario||"",
        "P2 Servicio":s.q2_servicio||"","Observación P2":s.observacion_q2||"","P3 Área técnica":s.q3_tecnica||"","Observación P3":s.observacion_q3||"",
        "P4 Área administrativa":s.q4_administrativa||"","Observación P4":s.observacion_q4||"","P5 Agilidad":s.q5_agilidad||"","P6 Recomendaría":s.q6_recomendaria||"","P7 Recomendación / felicitación":s.q7_recomendacion||""
      };});
      const ws=window.XLSX.utils.json_to_sheet(detail.length?detail:[{"Fecha":"","Asesor":"","Código y nombre del usuario":""}],{});
      ws["!cols"]=[{wch:12},{wch:24},{wch:18},{wch:32},{wch:18},{wch:32},{wch:18},{wch:32},{wch:22},{wch:32},{wch:28},{wch:18},{wch:45}];
      window.XLSX.utils.book_append_sheet(wb,ws,"Detalle de encuestas");

      const charts=[
        {type:"pie",title:"¿Recomendaría nuestros servicios?",seriesName:"Recomendación",catRef:excelRangeRef(sheetName,q6Row,q6Row+1,0),catCache:["SI","NO"],valRef:excelRangeRef(sheetName,q6Row,q6Row+1,1),valCache:[yes,no],anchor:{fromCol:5,fromRow:2,toCol:12,toRow:18}},
        {type:"bar",title:"Pregunta 2 · Servicio",seriesName:"Cantidad",catRef:excelRangeRef(sheetName,q2Row,q2Row+q2Opts.length-1,0),catCache:q2Opts,valRef:excelRangeRef(sheetName,q2Row,q2Row+q2Opts.length-1,1),valCache:distCounts("q2_servicio",q2Opts),anchor:{fromCol:5,fromRow:19,toCol:12,toRow:35}}
      ];
      await saveWorkbookWithCharts(wb,1,charts,`reporte-encuestas-${getTodayISO()}.xlsx`);
      showToast("Excel de encuestas descargado correctamente.");
    }catch(e){console.error(e);showToast("No fue posible generar el Excel de encuestas.",true);}
  }

  function renderUsers(){const tbody=id("tabla-usuarios");if(!tbody)return;tbody.innerHTML=advisors.map(a=>{const name=[a.nombre,a.apellido].filter(Boolean).join(" ")||"—";return `<tr><td><strong>${escapeHTML(name)}</strong></td><td>${escapeHTML(a.email||"—")}</td><td>${escapeHTML(a.zona||"—")}</td><td>${Number(a.meta_mensual)||50}</td><td>${a.activo===false?'<span class="badge badge-disabled">Inhabilitado</span>':'<span class="badge badge-active">Activo</span>'}</td><td class="action-cell"><button class="btn-small" onclick="editAdvisor('${a.id}')">Editar</button><button class="btn-small" onclick="toggleAdvisor('${a.id}',${a.activo!==false})">${a.activo===false?"Habilitar":"Inhabilitar"}</button><button class="btn-delete" onclick="deleteAdvisor('${a.id}')">Eliminar</button></td></tr>`;}).join("")||'<tr class="empty-row"><td colspan="6">No hay asesores registrados.</td></tr>';}

  async function saveAdminUser(e){e.preventDefault();const idUser=id("admin-user-id").value;const body={nombre:value("admin-user-nombre"),apellido:value("admin-user-apellido"),documento:value("admin-user-documento"),telefono:value("admin-user-telefono"),zona:value("admin-user-zona"),email:value("admin-user-email"),meta_mensual:Math.max(1,Number(id("admin-user-meta").value)||50)};if(!idUser){const password=id("admin-user-password").value;if(password.length<6){showToast("La contraseña debe tener mínimo 6 caracteres.",true);return;}const {data:{session}}=await sbClient.auth.getSession();if(!session){showToast("Sesión no disponible.",true);return;}const result=await fetchAdminFunction("create",{...body,password});if(result.error){showToast(result.error,true);return;}showToast("Asesor creado correctamente.");resetUserForm();await loadAdminData();return;}const result=await fetchAdminFunction("update",{user_id:idUser,...body});if(result.error){showToast(result.error,true);return;}showToast("Asesor actualizado.");resetUserForm();await loadAdminData();}
  async function fetchAdminFunction(action,payload){const {data:{session}}=await sbClient.auth.getSession();if(!session)return{error:"Sesión no disponible."};try{const r=await fetch(`${SUPABASE_URL}/functions/v1/admin-users`,{method:"POST",headers:{"Content-Type":"application/json",Authorization:`Bearer ${session.access_token}`},body:JSON.stringify({action,...payload})});const j=await r.json().catch(()=>({}));return r.ok?{data:j}:{error:j.error||`Error ${r.status}`};}catch(e){return{error:"No se pudo contactar la función de administración. Debes desplegar supabase/functions/admin-users."};}}
  function editAdvisor(uid){const a=advisors.find(x=>x.id===uid);if(!a)return;id("admin-user-id").value=a.id;["nombre","apellido","documento","telefono","zona","email"].forEach(k=>id(`admin-user-${k}`).value=a[k]||"");id("admin-user-meta").value=Number(a.meta_mensual)||50;id("admin-user-password").value="";id("btn-save-user").textContent="Actualizar asesor";id("btn-cancel-user-edit").classList.remove("hidden");document.getElementById("vista-usuarios").scrollIntoView({behavior:"smooth"});}
  async function toggleAdvisor(uid,active){const {error}=await sbClient.from("perfiles").update({activo:!active}).eq("id",uid);if(error){showToast(error.message,true);return;}showToast(active?"Asesor inhabilitado.":"Asesor habilitado.");await loadAdminData();}
  async function deleteAdvisor(uid){const a=advisors.find(x=>x.id===uid);if(!a)return;if(!confirm(`¿Eliminar a ${[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email}? Solo se podrá eliminar si no tiene operaciones registradas.`))return;const result=await fetchAdminFunction("delete",{user_id:uid});if(result.error){showToast(result.error,true);return;}showToast("Asesor eliminado.");await loadAdminData();}
  function resetUserForm(){id("admin-user-form").reset();id("admin-user-id").value="";id("admin-user-meta").value=50;id("btn-save-user").textContent="Crear asesor";id("btn-cancel-user-edit").classList.add("hidden");}

  async function saveConfig(e){e.preventDefault();let logo=config.logo_url||"";const file=id("config-logo").files[0];if(file){if(file.size>2*1024*1024){showToast("La imagen debe pesar máximo 2 MB.",true);return;}const ext=(file.name.split(".").pop()||"png").toLowerCase();const path=`tvmax/logo-${Date.now()}.${ext}`;const up=await sbClient.storage.from("app-assets").upload(path,file,{cacheControl:"31536000",upsert:false,contentType:file.type});if(up.error){showToast(`No fue posible almacenar el logo: ${up.error.message}`,true);return;}logo=sbClient.storage.from("app-assets").getPublicUrl(path).data.publicUrl;}const color=id("config-color").value;const {error}=await sbClient.from("configuracion").upsert({id:1,color_principal:color,logo_url:logo,updated_by:currentUser.id},{onConflict:"id"});if(error){showToast(error.message,true);return;}config={color_principal:color,logo_url:logo};cacheSet("config:tvmax",config);applyTheme();renderConfig();showToast("Configuración guardada. El logo ahora usa almacenamiento y caché del navegador.");}
  async function removeLogo(){const {error}=await sbClient.from("configuracion").upsert({id:1,color_principal:config.color_principal,logo_url:"",updated_by:currentUser.id},{onConflict:"id"});if(error){showToast(error.message,true);return;}cacheInvalidate("config:tvmax");config.logo_url="";renderConfig();showToast("Imagen retirada del reporte.");}
  function renderConfig(){id("config-color").value=config.color_principal||"#8b5cf6";id("logo-preview").innerHTML=config.logo_url?`<img loading="lazy" src="${config.logo_url}" alt="Logo de empresa">`:'<span>LOGO</span>';}
  function applyTheme(){document.documentElement.style.setProperty("--purple-primary",config.color_principal||"#8b5cf6");}
  function clearAdminFilters(){adminReportSales=null;["filtroAdminTexto","filtroEstadoAdmin","filtroTipoAdmin","filtroServicioAdmin","filtroZonaAdmin","filtroDesdeAdmin","filtroHastaAdmin"].forEach(x=>id(x).value="");selectedAdvisorIds=[];syncAdvisorFilterUI();renderAdmin();}

  function getCurrentMonthKey(){
    const now=new Date();
    return `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`;
  }
  function getMonthLabel(){
    return new Date().toLocaleDateString("es-CO",{month:"long",year:"numeric"});
  }
  function getAdminMonthlyGoal(){
    const people=adminGoalAdvisors.length?adminGoalAdvisors:advisors;
    return people.filter(a=>a.activo!==false).reduce((sum,a)=>sum+Math.max(1,Number(a.meta)||Number(a.meta_mensual)||50),0);
  }
  function getAdminMonthlySales(){
    return Array.isArray(adminMonthlySales)?adminMonthlySales.filter(s=>s.tipo_operacion==="Venta"):[];
  }
  function updateAdminDashboard(dashboard=null){
    const d=dashboard||{};
    const total=Number(d.total)||0,ventas=Number(d.ventas)||0,recon=Number(d.reconexiones)||0,p=Number(d.pendientes)||0,r=Number(d.realizadas)||0,c=Number(d.canceladas)||0;
    const dashboardAdvisors=Array.isArray(d.asesores)?d.asesores:adminGoalAdvisors;
    setText("dash-total",total);setText("dash-ventas",ventas);setText("dash-reconexiones",recon);setText("dash-pendientes",p);setText("dash-realizadas",r);setText("dash-canceladas",c);
    // IMPORTANTE: la meta del administrador cuenta SOLO ventas.
    const made=getAdminMonthlySales().length||ventas,goal=getAdminMonthlyGoal(),pct=goal?Math.min(100,Math.round(made/goal*100)):0;
    setText("dash-admin-goal-title",`${made} / ${goal} ventas`);
    setText("dash-admin-goal-period",`Meta total de ${dashboardAdvisors.length} asesor${dashboardAdvisors.length===1?"":"es"} activos para ${getMonthLabel()}. El avance individual incluye ventas y reconexiones del mes.`);
    if(id("dash-admin-goal-bar"))id("dash-admin-goal-bar").style.width=`${pct}%`;
    setText("dash-admin-goal-percent",`${pct}%`);
    setText("dash-admin-goal-detail",`${made} ventas realizadas de ${goal}`);
    if(id("dash-admin-goal-breakdown"))id("dash-admin-goal-breakdown").innerHTML=dashboardAdvisors.filter(a=>a.activo!==false).map(a=>{
      const n=[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email||"Asesor",g=Math.max(1,Number(a.meta)||50),count=Number(a.realizadas)||0,ap=Math.min(100,Math.round(count/g*100));
      return `<div class="admin-goal-breakdown-row"><span>${escapeHTML(n)}</span><strong>${count}/${g}</strong><small>${ap}%</small></div>`;
    }).join("")||'<span class="muted">No hay asesores activos.</span>';
    id("dash-goals-list").innerHTML=dashboardAdvisors.filter(a=>a.activo!==false).map(a=>{
      const n=[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email,count=Number(a.realizadas)||0,g=Math.max(1,Number(a.meta)||50),ap=Math.min(100,Math.round(count/g*100));
      return `<div class="goal-list-row"><div><strong>${escapeHTML(n)}</strong><small>${count} / ${g} operaciones</small></div><div class="mini-progress"><span style="width:${ap}%"></span></div><b>${ap}%</b></div>`;
    }).join("")||'<p class="muted">No hay asesores activos.</p>';
    const serviceCounts=d.servicios&&typeof d.servicios==="object"?d.servicios:{};const counts=SERVICES.map(s=>({s,n:Number(serviceCounts[s])||0}));const max=Math.max(1,...counts.map(x=>x.n));id("dash-services-list").innerHTML=counts.map(x=>`<div class="mini-bar-row"><span>${x.s}</span><div><i style="width:${x.n/max*100}%"></i></div><strong>${x.n}</strong></div>`).join("");
  }

  async function loadAdminSalesForFilters(force=false){
    if(!currentUser||currentProfile?.rol!=="administrador")return;
    const from=value("filtroDesdeAdmin"),to=value("filtroHastaAdmin");
    if(!from&&!to){adminReportSales=null;renderAdmin();return;}
    const key=`sales-report:tvmax:${from||"all"}:${to||"all"}`;
    if(!force&&adminReportSales?.__key===key)return;
    let q=sbClient.from("ventas").select("id,asesor_id,tipo_operacion,codigo_cliente,codigo_servicio,descripcion_servicio,zona,fecha_venta,estado_instalacion,fecha_instalacion,created_at,updated_at,servicio,perfiles:asesor_id(id,nombre,apellido,zona,email,meta_mensual,activo)").order("fecha_venta",{ascending:false}).order("id",{ascending:false});
    if(from)q=q.gte("fecha_venta",from); if(to)q=q.lte("fecha_venta",to);
    const r=await q; if(r.error)throw r.error; adminReportSales=r.data||[];
    adminReportSales.__key=key; renderAdmin();
  }
  async function loadAdvisorSalesForFilters(force=false){
    if(!currentUser||currentProfile?.rol!=="asesor")return;
    const from=value("filtroAsesorDesde"),to=value("filtroAsesorHasta");
    if(!from&&!to){advisorReportSales=null;renderAdvisorTable();return;}
    const key=`sales-advisor-report:tvmax:${currentUser.id}:${from||"all"}:${to||"all"}`;
    if(!force&&advisorReportLoadedKey===key)return;
    const q=sbClient.from("ventas").select("id,asesor_id,tipo_operacion,codigo_cliente,codigo_servicio,descripcion_servicio,zona,fecha_venta,estado_instalacion,fecha_instalacion,created_at,updated_at,servicio").eq("asesor_id",currentUser.id).order("fecha_venta",{ascending:false}).order("id",{ascending:false});
    const r=from?await q.gte("fecha_venta",from).lte("fecha_venta",to||from):await q.lte("fecha_venta",to);
    if(r.error)throw r.error; advisorReportSales=r.data||[];advisorReportLoadedKey=key;renderAdvisorTable();
  }

  async function loadSurveyReportData(force=false){
    if(!currentUser)return;
    const from=value("filtroEncuestaDesde")||monthStartISO();
    const to=value("filtroEncuestaHasta")||todayISO();
    const key=`survey-report:tvmax:${from}:${to}`;
    if(!force&&surveyReportLoadedKey===key&&surveys.length)return;
    if(surveyReportLoading&&!force)return surveyReportLoading;
    surveyReportLoading=(async()=>{
      const r=await sbClient.from("encuestas").select("id,asesor_id,codigo_nombre_usuario,q2_servicio,observacion_q2,q3_tecnica,observacion_q3,q4_administrativa,observacion_q4,q5_agilidad,observacion_q5,q6_recomendaria,observacion_q6,q7_recomendacion,fecha_encuesta,created_at,updated_at,perfiles:asesor_id(id,nombre,apellido,email,zona,activo,rol)").gte("fecha_encuesta",from).lte("fecha_encuesta",to).order("id",{ascending:false});
      if(r.error)throw r.error;
      surveys=r.data||[];surveyReportLoadedKey=key;populateSurveyZoneFilter();populateSurveyAdvisorFilter();renderSurveyReport();
    })().finally(()=>{surveyReportLoading=null;});
    return surveyReportLoading;
  }

  function ensureSurveyZoneFilter(){
    if(id("filtroEncuestaZona"))return;
    const base=id("filtroEncuestaAsesor");if(!base)return;
    const group=base.closest(".form-group");if(!group||!group.parentElement)return;
    const wrapper=document.createElement("div");wrapper.className="form-group";wrapper.innerHTML='<label for="filtroEncuestaZona">Zona</label><select id="filtroEncuestaZona"><option value="">Todas las zonas</option></select>';
    group.parentElement.insertBefore(wrapper,group.nextSibling);
    const z=id("filtroEncuestaZona");z.addEventListener("change",()=>renderSurveyReport());
    const table=id("tabla-reporte-encuestas");const head=table?.closest("table")?.querySelector("thead tr");
    if(head&&!head.querySelector("[data-zone-column]")){const th=document.createElement("th");th.textContent="Zona";th.setAttribute("data-zone-column","1");head.insertBefore(th,head.children[2]||null);}
    populateSurveyZoneFilter();
  }
  function populateSurveyZoneFilter(){
    const el=id("filtroEncuestaZona");if(!el)return;const selected=el.value;
    const zones=[...new Set([...surveys.map(s=>s.perfiles?.zona),...advisors.map(a=>a.zona)].filter(Boolean))].sort((a,b)=>a.localeCompare(b));
    el.innerHTML='<option value="">Todas las zonas</option>'+zones.map(z=>`<option value="${escapeHTML(z)}">${escapeHTML(z)}</option>`).join("");el.value=selected;
  }

  function buildReportHTML(){const filtered=getFilteredAdminSales(),total=filtered.length,ventas=filtered.filter(s=>s.tipo_operacion==="Venta").length,recon=filtered.filter(s=>s.tipo_operacion==="Reconexión").length,otros=filtered.filter(s=>s.tipo_operacion==="Otros").length,real=filtered.filter(s=>s.estado_instalacion==="REALIZADA").length,pending=filtered.filter(s=>s.estado_instalacion==="PENDIENTE").length,cancel=filtered.filter(s=>s.estado_instalacion==="CANCELADA").length,pct=n=>total?Math.round(n/total*100):0;
    const advisorMap={};filtered.forEach(s=>{const a=s.perfiles||{},n=[a.nombre,a.apellido].filter(Boolean).join(" ")||"Sin asesor";if(!advisorMap[n])advisorMap[n]={ventas:0,meta:Math.max(1,Number(a.meta_mensual)||50)};if(isGoalOperation(s))advisorMap[n].ventas++;});const advisorRows=Object.entries(advisorMap).sort((a,b)=>b[1].ventas-a[1].ventas).map(([n,d])=>{const gp=Math.min(100,Math.round(d.ventas/d.meta*100));return `<div class="print-advisor-row"><div class="print-advisor-label"><span>${escapeHTML(n)}</span><strong>${d.ventas}/${d.meta} operaciones · ${gp}%</strong></div><div class="print-bar-track"><div class="print-bar-fill" style="width:${gp}%"></div></div></div>`;}).join("")||'<div class="print-empty-chart">Sin datos</div>';
    const adminMonthlySales=getAdminMonthlySales(),adminGoal=getAdminMonthlyGoal(),adminMade=adminMonthlySales.length,adminPct=adminGoal?Math.min(100,Math.round(adminMade/adminGoal*100)):0;
    const goalPeople=adminGoalAdvisors.length?adminGoalAdvisors:advisors;
    const adminGoalRows=goalPeople.filter(a=>a.activo!==false).map(a=>{const n=[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email||"Asesor",g=Math.max(1,Number(a.meta)||Number(a.meta_mensual)||50),made=adminMonthlySales.filter(s=>s.asesor_id===a.id).length,p=Math.min(100,Math.round(made/g*100));return `<tr><td>${escapeHTML(n)}</td><td>${made}</td><td>${g}</td><td>${p}%</td></tr>`;}).join("")||'<tr><td colspan="4" class="print-empty-row">No hay asesores activos.</td></tr>';
    const desde=value("filtroDesdeAdmin"),hasta=value("filtroHastaAdmin"),period=desde||hasta?`${desde?formatDate(desde):"Inicio"} – ${hasta?formatDate(hasta):"Actual"}`:"Todos los periodos";const rows=filtered.map(s=>{const a=s.perfiles||{},n=[a.nombre,a.apellido].filter(Boolean).join(" ")||"—";return `<tr><td>${escapeHTML(n)}</td><td>${escapeHTML(s.tipo_operacion||"—")}</td><td>${escapeHTML([s.servicio,s.descripcion_servicio].filter(Boolean).join(" · "))}</td><td>${escapeHTML(s.zona||"—")}</td><td>${escapeHTML(statusLabel(s.estado_instalacion))}</td></tr>`}).join("");
    return `<div class="print-report-sheet">${config.logo_url?`<div class="print-logo"><img loading="lazy" src="${config.logo_url}" alt="Logo"></div>`:""}<div class="print-header"><div><span class="print-kicker">REPORTE DE OPERACIONES</span><h1>Ventas e instalaciones</h1><p>Periodo: <strong>${escapeHTML(period)}</strong></p></div><div class="print-generated">Generado: ${new Date().toLocaleString("es-CO")}</div></div><div class="print-summary"><div class="print-summary-card"><span>Total operaciones</span><strong>${total}</strong></div><div class="print-summary-card"><span>Ventas</span><strong>${ventas}</strong></div><div class="print-summary-card"><span>Reconexiones</span><strong>${recon}</strong></div><div class="print-summary-card"><span>Realizadas</span><strong>${real}</strong></div><div class="print-summary-card"><span>Pendientes</span><strong>${pending}</strong></div><div class="print-summary-card"><span>Canceladas</span><strong>${cancel}</strong></div><div class="print-summary-card print-admin-goal-card"><span>Meta administrador · ${escapeHTML(getMonthLabel())}</span><strong>${adminMade}/${adminGoal} · ${adminPct}%</strong><small>Suma de las metas de ${goalPeople.filter(a=>a.activo!==false).length} asesor${goalPeople.filter(a=>a.activo!==false).length===1?"":"es"} activos · solo ventas</small></div></div><section class="print-charts"><div class="print-chart-card"><h2>Operaciones</h2><div class="print-donut">${donutSVG(pct(ventas))}<div class="print-donut-center"><strong>${total}</strong><span>total</span></div></div><div class="print-legend"><span>Venta <strong>${pct(ventas)}%</strong></span><span>Reconexión <strong>${pct(recon)}%</strong></span><span>Otros <strong>${pct(otros)}%</strong></span></div></div><div class="print-chart-card"><h2>Estado</h2><div class="print-donut">${donutSVG(pct(real))}<div class="print-donut-center"><strong>${pct(real)}%</strong><span>realizadas</span></div></div><div class="print-legend"><span>Realizada <strong>${pct(real)}%</strong></span><span>Pendiente <strong>${pct(pending)}%</strong></span><span>Cancelada <strong>${pct(cancel)}%</strong></span></div></div><div class="print-chart-card print-advisor-chart"><h2>Cumplimiento de meta por asesor</h2>${advisorRows}</div></section><section class="print-table-section"><div class="print-table-title"><div><span class="print-kicker">META MENSUAL</span><h2>Meta del administrador</h2></div><strong>${adminMade}/${adminGoal} · ${adminPct}%</strong></div><div class="print-admin-goal-bar"><div><span style="width:${adminPct}%"></span></div></div><p class="print-meta-note">La meta del administrador corresponde a la suma de las metas mensuales de los asesores activos. El avance utiliza únicamente las ventas de ${escapeHTML(getMonthLabel())}; al comenzar un nuevo mes, el contador vuelve a cero.</p><div class="print-table-scroll"><table><thead><tr><th>Asesor</th><th>Operaciones del mes</th><th>Meta mensual</th><th>% cumplimiento</th></tr></thead><tbody>${adminGoalRows}</tbody></table></div></section><section class="print-table-section"><div class="print-table-title"><div><span class="print-kicker">DETALLE</span><h2>Operaciones registradas</h2></div><strong>${total} resultado${total===1?"":"s"}</strong></div><div class="print-table-scroll"><table><thead><tr><th>Asesor</th><th>Operación</th><th>Servicio</th><th>Zona</th><th>Estado</th></tr></thead><tbody>${rows||'<tr><td colspan="5" class="print-empty-row">No hay registros.</td></tr>'}</tbody></table></div></section></div>`;
  }
  function buildReportSummaryHTML(){
    const filtered=getFilteredAdminSales(),total=filtered.length,ventas=filtered.filter(s=>s.tipo_operacion==="Venta").length,recon=filtered.filter(s=>s.tipo_operacion==="Reconexión").length,otros=filtered.filter(s=>s.tipo_operacion==="Otros").length,real=filtered.filter(s=>s.estado_instalacion==="REALIZADA").length,pending=filtered.filter(s=>s.estado_instalacion==="PENDIENTE").length,cancel=filtered.filter(s=>s.estado_instalacion==="CANCELADA").length,pct=n=>total?Math.round(n/total*100):0;
    const adminMonthlySales=getAdminMonthlySales(),adminGoal=getAdminMonthlyGoal(),adminMade=adminMonthlySales.length,adminPct=adminGoal?Math.min(100,Math.round(adminMade/adminGoal*100)):0;
    const goalPeople=adminGoalAdvisors.length?adminGoalAdvisors:advisors;
    const adminGoalRows=goalPeople.filter(a=>a.activo!==false).map(a=>{const n=[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email||"Asesor",g=Math.max(1,Number(a.meta)||Number(a.meta_mensual)||50),made=adminMonthlySales.filter(s=>s.asesor_id===a.id).length,p=Math.min(100,Math.round(made/g*100));return `<tr><td>${escapeHTML(n)}</td><td>${made}</td><td>${g}</td><td>${p}%</td></tr>`;}).join("")||'<tr><td colspan="4" class="print-empty-row">No hay asesores activos.</td></tr>';
    const desde=value("filtroDesdeAdmin"),hasta=value("filtroHastaAdmin"),period=desde||hasta?`${desde?formatDate(desde):"Inicio"} – ${hasta?formatDate(hasta):"Actual"}`:"Todos los periodos";
    return `<div class="print-report-sheet compact-pdf">${config.logo_url?`<div class="print-logo"><img loading="lazy" src="${config.logo_url}" alt="Logo"></div>`:""}<div class="print-header"><div><span class="print-kicker">REPORTE DE OPERACIONES</span><h1>Ventas e instalaciones</h1><p>Periodo: <strong>${escapeHTML(period)}</strong></p></div><div class="print-generated">Generado: ${new Date().toLocaleString("es-CO")}</div></div><div class="print-summary"><div class="print-summary-card"><span>Total operaciones</span><strong>${total}</strong></div><div class="print-summary-card"><span>Ventas</span><strong>${ventas}</strong></div><div class="print-summary-card"><span>Reconexiones</span><strong>${recon}</strong></div><div class="print-summary-card"><span>Realizadas</span><strong>${real}</strong></div><div class="print-summary-card"><span>Pendientes</span><strong>${pending}</strong></div><div class="print-summary-card"><span>Canceladas</span><strong>${cancel}</strong></div><div class="print-summary-card print-admin-goal-card"><span>Meta administrador · ${escapeHTML(getMonthLabel())}</span><strong>${adminMade}/${adminGoal} · ${adminPct}%</strong><small>Solo ventas</small></div></div><section class="print-charts"><div class="print-chart-card"><h2>Operaciones</h2><div class="print-donut">${donutSVG(pct(ventas))}<div class="print-donut-center"><strong>${total}</strong><span>total</span></div></div><div class="print-legend"><span>Venta <strong>${pct(ventas)}%</strong></span><span>Reconexión <strong>${pct(recon)}%</strong></span><span>Otros <strong>${pct(otros)}%</strong></span></div></div><div class="print-chart-card"><h2>Estado</h2><div class="print-donut">${donutSVG(pct(real))}<div class="print-donut-center"><strong>${pct(real)}%</strong><span>realizadas</span></div></div><div class="print-legend"><span>Realizada <strong>${pct(real)}%</strong></span><span>Pendiente <strong>${pct(pending)}%</strong></span><span>Cancelada <strong>${pct(cancel)}%</strong></span></div></div></section><section class="print-table-section"><div class="print-table-title"><div><span class="print-kicker">META MENSUAL</span><h2>Meta del administrador</h2></div><strong>${adminMade}/${adminGoal} · ${adminPct}%</strong></div><div class="print-admin-goal-bar"><div><span style="width:${adminPct}%"></span></div></div><div class="print-table-scroll"><table><thead><tr><th>Asesor</th><th>Operaciones del mes</th><th>Meta mensual</th><th>% cumplimiento</th></tr></thead><tbody>${adminGoalRows}</tbody></table></div></section><p class="print-footnote">Resumen ejecutivo · el detalle completo de las ${total} operaciones registradas y los gráficos ampliados están disponibles en el archivo Excel.</p></div>`;
  }
  function buildSurveySummaryHTML(){
    const list=getFilteredSurveys(),total=list.length,yes=list.filter(s=>s.q6_recomendaria==="SI").length,no=list.filter(s=>s.q6_recomendaria==="NO").length;
    const pct=n=>total?Math.round(n/total*100):0;
    const advisorRows=surveyReportPeople().map(a=>{
      const rows=list.filter(s=>s.asesor_id===a.id), y=rows.filter(s=>s.q6_recomendaria==="SI").length, p=rows.length?Math.round(y/rows.length*100):0;
      if(!rows.length)return "";
      const name=[a.nombre,a.apellido].filter(Boolean).join(" ")||a.email||"Asesor";
      return `<tr><td>${escapeHTML(name)}</td><td>${rows.length}</td><td>${y}</td><td>${rows.length-y}</td><td>${p}%</td></tr>`;
    }).join("")||'<tr><td colspan="5" class="print-empty-row">No hay datos.</td></tr>';
    const dist=(field,opts)=>opts.map(o=>`<tr><td>${escapeHTML(o)}</td><td>${list.filter(s=>s[field]===o).length}</td><td>${pct(list.filter(s=>s[field]===o).length)}%</td></tr>`).join("");
    const period=value("filtroEncuestaDesde")||value("filtroEncuestaHasta")?`${value("filtroEncuestaDesde")?formatDate(value("filtroEncuestaDesde")):"Inicio"} – ${value("filtroEncuestaHasta")?formatDate(value("filtroEncuestaHasta")):"Actual"}`:"Todos los periodos";
    return `<div class="print-report-sheet survey-print-sheet compact-pdf">${config.logo_url?`<div class="print-logo"><img loading="lazy" src="${config.logo_url}" alt="Logo"></div>`:""}<div class="print-header"><div><span class="print-kicker">REPORTE DE ENCUESTAS</span><h1>Satisfacción de usuarios</h1><p>Periodo: <strong>${escapeHTML(period)}</strong></p></div><div class="print-generated">Generado: ${new Date().toLocaleString("es-CO")}</div></div><div class="print-summary"><div class="print-summary-card"><span>Total encuestas</span><strong>${total}</strong></div><div class="print-summary-card"><span>Recomiendan</span><strong>${yes} (${pct(yes)}%)</strong></div><div class="print-summary-card"><span>No recomiendan</span><strong>${no} (${pct(no)}%)</strong></div></div><section class="print-table-section"><div class="print-table-title"><div><span class="print-kicker">POR ASESOR</span><h2>Encuestas registradas por asesor</h2></div></div><div class="print-table-scroll"><table><thead><tr><th>Asesor</th><th>Total</th><th>Sí</th><th>No</th><th>% Sí</th></tr></thead><tbody>${advisorRows}</tbody></table></div></section><section class="print-table-section"><div class="print-table-title"><div><span class="print-kicker">POR ZONA</span><h2>Encuestas realizadas por zona</h2></div></div><div class="print-table-scroll"><table><thead><tr><th>Zona</th><th>Encuestas</th><th>%</th></tr></thead><tbody>${Object.entries(list.reduce((m,s)=>{const z=s.perfiles?.zona||"Sin zona";m[z]=(m[z]||0)+1;return m;},{})).sort((a,b)=>b[1]-a[1]).map(([z,n])=>`<tr><td>${escapeHTML(z)}</td><td>${n}</td><td>${pct(n)}%</td></tr>`).join("")||'<tr><td colspan="3" class="print-empty-row">No hay datos por zona.</td></tr>'}</tbody></table></div></section><section class="print-table-section"><div class="print-table-title"><div><span class="print-kicker">DISTRIBUCIÓN</span><h2>Respuestas por pregunta</h2></div></div><div class="survey-print-distributions"><div><h3>Pregunta 2</h3><p class="q-text">${escapeHTML(SURVEY_QUESTIONS.q2_servicio)}</p><table><thead><tr><th>Respuesta</th><th>Cantidad</th><th>%</th></tr></thead><tbody>${dist("q2_servicio",["BUENO","EXCELENTE","MALO","MUY MALO","REGULAR"])}</tbody></table></div><div><h3>Pregunta 3</h3><p class="q-text">${escapeHTML(SURVEY_QUESTIONS.q3_tecnica)}</p><table><thead><tr><th>Respuesta</th><th>Cantidad</th><th>%</th></tr></thead><tbody>${dist("q3_tecnica",["BUENO","EXCELENTE","MALO","MUY MALO","REGULAR"])}</tbody></table></div><div><h3>Pregunta 4</h3><p class="q-text">${escapeHTML(SURVEY_QUESTIONS.q4_administrativa)}</p><table><thead><tr><th>Respuesta</th><th>Cantidad</th><th>%</th></tr></thead><tbody>${dist("q4_administrativa",["BUENO","EXCELENTE","MALO","MUY MALO","REGULAR"])}</tbody></table></div><div><h3>Pregunta 5</h3><p class="q-text">${escapeHTML(SURVEY_QUESTIONS.q5_agilidad)}</p><table><thead><tr><th>Respuesta</th><th>Cantidad</th><th>%</th></tr></thead><tbody>${dist("q5_agilidad",["AGIL","DEMORADOS","MUY DEMORADOS","NI DEMORADOS NI AGIL"])}</tbody></table></div><div><h3>Pregunta 6</h3><p class="q-text">${escapeHTML(SURVEY_QUESTIONS.q6_recomendaria)}</p><table><thead><tr><th>Respuesta</th><th>Cantidad</th><th>%</th></tr></thead><tbody>${dist("q6_recomendaria",["SI","NO"])}</tbody></table></div></div></section><p class="print-footnote">Resumen ejecutivo · el detalle completo de las ${total} encuestas está disponible en el archivo Excel.</p></div>`;
  }

  function previewReport(builder=buildReportHTML){const modal=id("report-preview-modal"),content=id("report-preview-content");if(!modal||!content)return;content.innerHTML=builder();modal.classList.remove("hidden");modal.setAttribute("aria-hidden","false");document.body.classList.add("report-preview-open");}
  function closeReportPreview(){const modal=id("report-preview-modal");if(!modal)return;modal.classList.add("hidden");modal.setAttribute("aria-hidden","true");document.body.classList.remove("report-preview-open");}
  function printReport(builder=buildReportHTML){id("print-report").innerHTML=builder();window.print();}
  async function downloadPDF(builder=buildReportHTML,filePrefix="reporte-ventas"){
    const area=id("print-report");area.innerHTML=builder();area.classList.add("pdf-rendering");
    try{
      const areaRect=area.getBoundingClientRect();
      const areaHeight=area.scrollHeight;
      const breakEls=area.querySelectorAll('tr,.print-summary-card,.print-chart-card,.print-advisor-row,.survey-print-distributions>div,.print-table-title');
      const breakpointsCss=new Set([areaHeight]);
      breakEls.forEach(el=>{
        const r=el.getBoundingClientRect();
        const bottom=r.bottom-areaRect.top;
        if(bottom>0&&bottom<areaHeight)breakpointsCss.add(bottom);
      });
      const sortedBreaksCss=[...breakpointsCss].sort((a,b)=>a-b);
      const canvas=await html2canvas(area,{scale:2,useCORS:true,backgroundColor:"#ffffff"});
      const {jsPDF}=window.jspdf;const pdf=new jsPDF({orientation:"landscape",unit:"mm",format:"a4"});
      const pageW=297,pageH=210,margin=8,imgW=pageW-margin*2;
      const scaleY=canvas.height/areaHeight;
      const breakpointsPx=sortedBreaksCss.map(v=>v*scaleY);
      const pxPerPage=canvas.width*(pageH-margin*2)/imgW;
      let sourceY=0,first=true;
      while(sourceY<canvas.height-0.5){
        const target=Math.min(canvas.height,sourceY+pxPerPage);
        let cut=null;
        for(const bp of breakpointsPx){ if(bp>sourceY+0.5&&bp<=target+0.5)cut=bp; if(bp>target+0.5)break; }
        if(cut===null||cut<=sourceY)cut=target;
        const h=Math.min(cut-sourceY,canvas.height-sourceY);
        if(h<=0)break;
        const pageCanvas=document.createElement("canvas");pageCanvas.width=canvas.width;pageCanvas.height=h;
        pageCanvas.getContext("2d").drawImage(canvas,0,sourceY,canvas.width,h,0,0,canvas.width,h);
        const pageImg=pageCanvas.toDataURL("image/jpeg",0.95);
        const hMm=h*imgW/canvas.width;
        if(!first)pdf.addPage();
        pdf.addImage(pageImg,"JPEG",margin,margin,imgW,hMm);
        first=false;sourceY=cut;
      }
      const now=new Date().toISOString().slice(0,10);pdf.save(`${filePrefix}-${now}.pdf`);showToast("PDF descargado correctamente.");
    }catch(e){console.error(e);showToast("No fue posible generar el PDF.",true);}
    finally{area.classList.remove("pdf-rendering");}
  }

  // Reporte de avance individual para el asesor (usa "sales", que ya viene filtrado a sus propias operaciones).
  function buildAdvisorReportHTML(compact=false){
    const list=getFilteredAdvisorSales(),total=list.length,ventas=list.filter(s=>s.tipo_operacion==="Venta").length,recon=list.filter(s=>s.tipo_operacion==="Reconexión").length,otros=list.filter(s=>s.tipo_operacion==="Otros").length,real=list.filter(s=>s.estado_instalacion==="REALIZADA").length,pending=list.filter(s=>s.estado_instalacion==="PENDIENTE").length,cancel=list.filter(s=>s.estado_instalacion==="CANCELADA").length,pct=n=>total?Math.round(n/total*100):0;
    const now=new Date(),monthlyAdvisor=advisorReportDashboard?.asesores?.find(a=>a.id===currentUser?.id),monthly=Number(monthlyAdvisor?.realizadas)||0,goal=Math.max(1,Number(currentProfile?.meta_mensual)||50),gp=Math.min(100,Math.round(monthly/goal*100));
    const nombre=[currentProfile?.nombre,currentProfile?.apellido].filter(Boolean).join(" ")||currentProfile?.email||"Asesor";
    const rows=list.map(s=>`<tr><td>${operationBadge(s.tipo_operacion)}</td><td>${escapeHTML([s.servicio,s.descripcion_servicio].filter(Boolean).join(" · "))}</td><td>${escapeHTML(s.zona||"—")}</td><td>${escapeHTML(formatDate(s.fecha_venta))}</td><td>${escapeHTML(statusLabel(s.estado_instalacion))}</td></tr>`).join("");
    return `<div class="print-report-sheet${compact?" compact-pdf":""}">${config.logo_url?`<div class="print-logo"><img loading="lazy" src="${config.logo_url}" alt="Logo"></div>`:""}<div class="print-header"><div><span class="print-kicker">REPORTE DE AVANCE</span><h1>${escapeHTML(nombre)}</h1><p>Zona: <strong>${escapeHTML(currentProfile?.zona||"—")}</strong></p></div><div class="print-generated">Generado: ${new Date().toLocaleString("es-CO")}</div></div><div class="print-summary"><div class="print-summary-card"><span>Total operaciones</span><strong>${total}</strong></div><div class="print-summary-card"><span>Ventas</span><strong>${ventas}</strong></div><div class="print-summary-card"><span>Reconexiones</span><strong>${recon}</strong></div><div class="print-summary-card"><span>Realizadas</span><strong>${real}</strong></div><div class="print-summary-card"><span>Pendientes</span><strong>${pending}</strong></div><div class="print-summary-card"><span>Canceladas</span><strong>${cancel}</strong></div></div><section class="print-charts"><div class="print-chart-card"><h2>Operaciones</h2><div class="print-donut">${donutSVG(pct(ventas))}<div class="print-donut-center"><strong>${total}</strong><span>total</span></div></div><div class="print-legend"><span>Venta <strong>${pct(ventas)}%</strong></span><span>Reconexión <strong>${pct(recon)}%</strong></span><span>Otros <strong>${pct(otros)}%</strong></span></div></div><div class="print-chart-card"><h2>Estado</h2><div class="print-donut">${donutSVG(pct(real))}<div class="print-donut-center"><strong>${pct(real)}%</strong><span>realizadas</span></div></div><div class="print-legend"><span>Realizada <strong>${pct(real)}%</strong></span><span>Pendiente <strong>${pct(pending)}%</strong></span><span>Cancelada <strong>${pct(cancel)}%</strong></span></div></div><div class="print-chart-card print-advisor-chart"><h2>Meta mensual</h2><div class="print-advisor-row"><div class="print-advisor-label"><span>${escapeHTML(now.toLocaleDateString("es-CO",{month:"long",year:"numeric"}))}</span><strong>${monthly}/${goal} operaciones · ${gp}%</strong></div><div class="print-bar-track"><div class="print-bar-fill" style="width:${gp}%"></div></div></div><p class="print-meta-note">Incluye ventas y reconexiones del mes.</p></div></section><section class="print-table-section"><div class="print-table-title"><div><span class="print-kicker">DETALLE</span><h2>Mis operaciones</h2></div><strong>${total} resultado${total===1?"":"s"}</strong></div><div class="print-table-scroll"><table><thead><tr><th>Operación</th><th>Servicio</th><th>Zona</th><th>Fecha</th><th>Estado</th></tr></thead><tbody>${rows||'<tr><td colspan="5" class="print-empty-row">No hay registros.</td></tr>'}</tbody></table></div></section></div>`;
  }
  async function ensureAdvisorReportData(){const from=value("filtroAsesorDesde"),to=value("filtroAsesorHasta");if(from||to)await loadAdvisorSalesForFilters(false);advisorReportDashboard=await loadMonthlyDashboard(false);}
  async function previewAdvisorReport(){await ensureAdvisorReportData();previewReport(buildAdvisorReportHTML);}
  async function printAdvisorReport(){await ensureAdvisorReportData();printReport(buildAdvisorReportHTML);}
  async function downloadAdvisorPDF(){await ensureAdvisorReportData();downloadPDF(()=>buildAdvisorReportHTML(true),"mi-reporte");}

  async function downloadAdvisorExcel(){
    try{
      if(!window.XLSX){showToast("No se pudo cargar el módulo de Excel.",true);return;}
      await ensureAdvisorReportData();
      const list=getFilteredAdvisorSales(), now=new Date(), monthlyDashboard=await loadMonthlyDashboard(false), monthlyAdvisor=monthlyDashboard?.asesores?.find(a=>a.id===currentUser?.id), monthly=Number(monthlyAdvisor?.realizadas)||0, goal=Math.max(1,Number(currentProfile?.meta_mensual)||50), gp=Math.min(100,Math.round(monthly/goal*100));
      const nombre=[currentProfile?.nombre,currentProfile?.apellido].filter(Boolean).join(" ")||currentProfile?.email||"Asesor";
      const summary=[["REPORTE DE AVANCE MENSUAL"],["Asesor",nombre],["Periodo del detalle",value("filtroAsesorDesde")||value("filtroAsesorHasta")?`${value("filtroAsesorDesde")?formatDate(value("filtroAsesorDesde")):"Inicio"} – ${value("filtroAsesorHasta")?formatDate(value("filtroAsesorHasta")):"Actual"}`:"Operaciones recientes"],["Meta mensual",goal],["Operaciones del mes (Venta + Reconexión)",monthly],["Avance",`${gp}%`],["Pendientes de meta",Math.max(0,goal-monthly)],["Ventas en detalle",list.filter(s=>s.tipo_operacion==="Venta").length],["Reconexiones en detalle",list.filter(s=>s.tipo_operacion==="Reconexión").length]];
      const detail=list.map(s=>({"Fecha":s.fecha_venta||"","Operación":s.tipo_operacion||"","Código cliente":s.codigo_cliente||"","Servicio":s.servicio||"","Descripción":s.descripcion_servicio||"","Zona":s.zona||"","Estado":statusLabel(s.estado_instalacion)}));
      const wb=window.XLSX.utils.book_new(), ws=window.XLSX.utils.aoa_to_sheet(summary), wd=window.XLSX.utils.json_to_sheet(detail.length?detail:[{"Fecha":"","Operación":"","Código cliente":"","Servicio":"","Descripción":"","Zona":"","Estado":""}]);
      ws["!cols"]=[{wch:38},{wch:28}]; wd["!cols"]=[{wch:14},{wch:18},{wch:18},{wch:18},{wch:42},{wch:22},{wch:16}];
      window.XLSX.utils.book_append_sheet(wb,ws,"Avance mensual");window.XLSX.utils.book_append_sheet(wb,wd,"Detalle operaciones");
      window.XLSX.writeFile(wb,`mi-avance-${new Date().toISOString().slice(0,10)}.xlsx`);showToast("Excel de avance descargado.");
    }catch(e){console.error(e);showToast("No fue posible generar el Excel de avance.",true);}
  }

  async function downloadExcel(){
    try{
      if(!window.XLSX){showToast("No se pudo cargar el módulo de Excel.",true);return;}
      const filtered=getFilteredAdminSales();
      const total=filtered.length, ventas=filtered.filter(s=>s.tipo_operacion==="Venta").length, recon=filtered.filter(s=>s.tipo_operacion==="Reconexión").length, otros=filtered.filter(s=>s.tipo_operacion==="Otros").length, real=filtered.filter(s=>s.estado_instalacion==="REALIZADA").length, pending=filtered.filter(s=>s.estado_instalacion==="PENDIENTE").length, cancel=filtered.filter(s=>s.estado_instalacion==="CANCELADA").length;
      const pct=n=>total?Math.round(n/total*100):0;
      const advisorMap={};filtered.forEach(s=>{const a=s.perfiles||{},n=[a.nombre,a.apellido].filter(Boolean).join(" ")||"Sin asesor";if(!advisorMap[n])advisorMap[n]={ventas:0,meta:Math.max(1,Number(a.meta_mensual)||50)};if(isGoalOperation(s))advisorMap[n].ventas++;});
      const detail=filtered.map(s=>{const a=s.perfiles||{};return {"Asesor":[a.nombre,a.apellido].filter(Boolean).join(" ")||"—","Operación":s.tipo_operacion||"—","Servicio":[s.servicio,s.descripcion_servicio].filter(Boolean).join(" · "),"Zona":s.zona||"—","Estado":statusLabel(s.estado_instalacion)}});
      const ws=window.XLSX.utils.json_to_sheet(detail.length?detail:[{"Asesor":"","Operación":"","Servicio":"","Zona":"","Estado":""}],{header:["Asesor","Operación","Servicio","Zona","Estado"]});
      ws["!cols"]=[{wch:25},{wch:16},{wch:48},{wch:20},{wch:16}];
      const sheetName="Resumen y gráficos";
      const summary=[];const push=r=>summary.push(r);
      push(["REPORTE DE OPERACIONES"]);
      push(["Periodo", value("filtroDesdeAdmin")||value("filtroHastaAdmin")?`${value("filtroDesdeAdmin")?formatDate(value("filtroDesdeAdmin")):"Inicio"} – ${value("filtroHastaAdmin")?formatDate(value("filtroHastaAdmin")):"Actual"}`:"Todos los periodos"]);push([]);
      push(["RESUMEN GENERAL"]);push(["Indicador","Cantidad","Porcentaje"]);
      push(["Total operaciones",total,"100%"]);push(["Ventas",ventas,`${pct(ventas)}%`]);push(["Reconexiones",recon,`${pct(recon)}%`]);push(["Otros",otros,`${pct(otros)}%`]);push(["Realizadas",real,`${pct(real)}%`]);push(["Pendientes",pending,`${pct(pending)}%`]);push(["Canceladas",cancel,`${pct(cancel)}%`]);push([]);
      push(["META DEL ADMINISTRADOR · MES ACTUAL"]);push(["Indicador","Valor"]);push(["Ventas del mes",getAdminMonthlySales().length]);push(["Meta total de asesores activos",getAdminMonthlyGoal()]);push(["Cumplimiento",`${getAdminMonthlyGoal()?Math.min(100,Math.round(getAdminMonthlySales().length/getAdminMonthlyGoal()*100)):0}%`]);push(["Asesores incluidos",(adminGoalAdvisors.length?adminGoalAdvisors:advisors).filter(a=>a.activo!==false).length]);push([]);
      push(["GRÁFICO · OPERACIONES"]);push(["Categoría","Cantidad","%"]);
      const opsRow=summary.length;push(["Venta",ventas,pct(ventas)]);push(["Reconexión",recon,pct(recon)]);push(["Otros",otros,pct(otros)]);push([]);
      push(["GRÁFICO · ESTADO"]);push(["Estado","Cantidad","%"]);
      const estadoRow=summary.length;push(["Realizada",real,pct(real)]);push(["Pendiente",pending,pct(pending)]);push(["Cancelada",cancel,pct(cancel)]);push([]);
      push(["GRÁFICO · CUMPLIMIENTO DE META POR ASESOR (VENTAS + RECONEXIONES)"]);push(["Asesor","Operaciones","Meta","% cumplimiento"]);
      const advisorRow=summary.length;
      const advisorEntries=Object.entries(advisorMap).sort((a,b)=>b[1].ventas-a[1].ventas);
      if(advisorEntries.length)advisorEntries.forEach(([n,d])=>{const gp=Math.min(100,Math.round(d.ventas/d.meta*100));push([n,d.ventas,d.meta,gp]);});else push(["Sin datos","","",""]);
      const advisorCount=advisorEntries.length||1;

      const wr=window.XLSX.utils.aoa_to_sheet(summary);wr["!cols"]=[{wch:34},{wch:16},{wch:16},{wch:18}];
      const wb=window.XLSX.utils.book_new();window.XLSX.utils.book_append_sheet(wb,wr,sheetName);window.XLSX.utils.book_append_sheet(wb,ws,"Detalle");

      const charts=[
        {type:"pie",title:"Operaciones por tipo",seriesName:"Operaciones",catRef:excelRangeRef(sheetName,opsRow,opsRow+2,0),catCache:["Venta","Reconexión","Otros"],valRef:excelRangeRef(sheetName,opsRow,opsRow+2,1),valCache:[ventas,recon,otros],anchor:{fromCol:5,fromRow:2,toCol:12,toRow:18}},
        {type:"pie",title:"Operaciones por estado",seriesName:"Estado",catRef:excelRangeRef(sheetName,estadoRow,estadoRow+2,0),catCache:["Realizada","Pendiente","Cancelada"],valRef:excelRangeRef(sheetName,estadoRow,estadoRow+2,1),valCache:[real,pending,cancel],anchor:{fromCol:5,fromRow:19,toCol:12,toRow:35}},
        {type:"bar",title:"Cumplimiento de meta por asesor (%)",seriesName:"% cumplimiento",catRef:excelRangeRef(sheetName,advisorRow,advisorRow+advisorCount-1,0),catCache:advisorEntries.length?advisorEntries.map(([n])=>n):["Sin datos"],valRef:excelRangeRef(sheetName,advisorRow,advisorRow+advisorCount-1,3),valCache:advisorEntries.length?advisorEntries.map(([n,d])=>Math.min(100,Math.round(d.ventas/d.meta*100))):[0],anchor:{fromCol:5,fromRow:36,toCol:14,toRow:36+Math.max(14,advisorCount+2)}}
      ];
      await saveWorkbookWithCharts(wb,1,charts,`reporte-ventas-${new Date().toISOString().slice(0,10)}.xlsx`);
      showToast("Excel descargado con resumen y gráficos.");
    }catch(e){console.error(e);showToast("No fue posible generar el Excel.",true);}
  }

  function showAuthView(){["auth-view","register-view","vista-asesor","admin-dashboard","vista-admin","vista-usuarios","vista-configuracion"].forEach(x=>id(x).classList.add("hidden"));id("auth-view").classList.remove("hidden");id("session-area").classList.add("hidden");id("btn-menu").classList.add("hidden");id("sidebar").classList.add("hidden");}
  function showView(viewId){["auth-view","register-view","vista-asesor","vista-encuestas","admin-dashboard","vista-admin","vista-reporte-encuestas","vista-usuarios","vista-configuracion","vista-respaldo"].forEach(x=>id(x).classList.add("hidden"));id(viewId).classList.remove("hidden");if(viewId!=="auth-view"&&currentProfile){id("session-area").classList.remove("hidden");id("btn-menu").classList.remove("hidden");id("sidebar").classList.remove("hidden");}if(viewId==="vista-admin")loadAdvisorsForFilters();
    if(viewId==="vista-reporte-encuestas"){ensureSurveyZoneFilter();loadSurveyReportData().catch(e=>console.warn("No fue posible cargar el reporte de encuestas",e));}
  }
  async function logout(){const {error}=await sbClient.auth.signOut();if(error)showToast("No fue posible cerrar la sesión.",true);}
  function installationStatus(s){if(s==="REALIZADA")return '<span class="badge badge-complete">Realizada</span>';if(s==="CANCELADA")return '<span class="badge badge-cancelled">Cancelada</span>';return '<span class="badge badge-pending">Pendiente</span>';}
  function statusLabel(s){return s==="REALIZADA"?"Realizada":s==="CANCELADA"?"Cancelada":"Pendiente";}
  function operationBadge(t){if(t==="Reconexión")return '<span class="badge badge-reconnection">Reconexión</span>';if(t==="Otros")return '<span class="badge badge-other">Otros</span>';return '<span class="badge badge-sale">Venta</span>';}
  function serviceBadge(s){return `<span class="badge badge-service">${escapeHTML(s||"Otros")}</span>`;}
  function formatDate(d){if(!d)return "—";const p=d.split("-");return p.length===3?`${p[2]}/${p[1]}/${p[0]}`:escapeHTML(d);}
  function setTodayDefault(){const x=id("fechaVenta");if(x&&!x.value)x.value=getTodayISO();}function getTodayISO(){const n=new Date(),o=n.getTimezoneOffset(),l=new Date(n.getTime()-o*60000);return l.toISOString().slice(0,10);}
  function debounce(fn,ms){let t;return(...args)=>{clearTimeout(t);t=setTimeout(()=>fn(...args),ms);};}
  function value(x){const el=id(x);return el&&typeof el.value==="string"?el.value.trim():"";}function id(x){return document.getElementById(x);}function setText(x,v){if(id(x))id(x).textContent=v;}
  function escapeHTML(v){return String(v??"").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;");}
  async function downloadBackup(){
    const btn=id("btn-download-backup"),status=id("backup-status");
    if(!window.XLSX){showToast("No se pudo cargar el módulo de Excel.",true);return;}
    setButtonBusy(btn,true,"Generando respaldo...");
    try{
      const [ventasRes,perfilesRes,encuestasRes]=await Promise.all([
        sbClient.from("ventas").select("*").order("id",{ascending:true}),
        sbClient.from("perfiles").select("*").order("created_at",{ascending:true}),
        sbClient.from("encuestas").select("*").order("id",{ascending:true})
      ]);
      if(ventasRes.error||perfilesRes.error||encuestasRes.error){console.error(ventasRes.error||perfilesRes.error||encuestasRes.error);showToast("No fue posible generar el respaldo.",true);return;}
      const wb=window.XLSX.utils.book_new();
      const wsVentas=window.XLSX.utils.json_to_sheet(ventasRes.data&&ventasRes.data.length?ventasRes.data:[{id:""}]);
      const wsPerfiles=window.XLSX.utils.json_to_sheet(perfilesRes.data&&perfilesRes.data.length?perfilesRes.data:[{id:""}]);
      const wsEncuestas=window.XLSX.utils.json_to_sheet(encuestasRes.data&&encuestasRes.data.length?encuestasRes.data:[{id:""}]);
      window.XLSX.utils.book_append_sheet(wb,wsVentas,"Ventas");
      window.XLSX.utils.book_append_sheet(wb,wsPerfiles,"Perfiles");
      window.XLSX.utils.book_append_sheet(wb,wsEncuestas,"Encuestas");
      const now=new Date();
      window.XLSX.writeFile(wb,`respaldo-cabletelco-${now.toISOString().slice(0,10)}.xlsx`);
      status.textContent=`Último respaldo generado: ${now.toLocaleString("es-CO")} · ${ventasRes.data.length} ventas, ${perfilesRes.data.length} perfiles, ${encuestasRes.data.length} encuestas.`;
      showToast("Respaldo generado correctamente.");
    }catch(e){console.error(e);showToast("No fue posible generar el respaldo.",true);}
    finally{setButtonBusy(btn,false,"⭳ Descargar respaldo completo");}
  }

  function setButtonBusy(b,busy,text){if(!b)return;b.disabled=busy;b.textContent=text;}function authError(e){const m=(e?.message||"").toLowerCase();if(m.includes("invalid login credentials"))return "Correo o contraseña incorrectos.";if(m.includes("email not confirmed"))return "Debes confirmar tu correo antes de iniciar sesión.";if(m.includes("user already registered"))return "Ese correo ya está registrado.";return e?.message||"No fue posible completar la operación.";}
  function fileToDataURL(file){return new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=rej;r.readAsDataURL(file);});}
  let toastTimer;function showToast(msg,error=false){const t=id("toast");t.textContent=msg;t.classList.toggle("error",error);t.classList.add("show");clearTimeout(toastTimer);toastTimer=setTimeout(()=>t.classList.remove("show"),3500);}

  window.setInstallation=setInstallation;window.deleteSale=deleteSale;window.editAdvisor=editAdvisor;window.toggleAdvisor=toggleAdvisor;window.deleteAdvisor=deleteAdvisor;
})();
