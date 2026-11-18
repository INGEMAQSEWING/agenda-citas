/**
 * Agenda de citas — programa del salón (Google Apps Script).
 * Va pegado dentro de una Hoja de cálculo de Google (Extensiones > Apps Script).
 * Guarda la configuración y las citas en esa hoja y envía los correos.
 */

// ====== LO ÚNICO QUE SE CAMBIA ======
const CLAVE_SALON = "CAMBIA-ESTA-CLAVE";   // clave para entrar al panel del salón
const APP_URL = "https://ingemaqsewing.github.io/agenda-citas/";  // dirección de la página
// ===================================

const ZONA = "America/Bogota";
const HOJA_CITAS = "Citas", HOJA_CONFIG = "Config";
const COLS = ["id","token","estado","fecha","hora","min","servicioId","servicio","icono","nombre","celular","correo","nota","motivo","creada","actualizada"];

function doGet(e){ return responder(manejar((e && e.parameter) || {})); }
function doPost(e){ let p = {}; try { p = JSON.parse(e.postData.contents); } catch(x){} return responder(manejar(p)); }
function responder(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function manejar(p){
  try {
    switch (p.accion){
      case "publico":      return {ok:true, cfg:leerCfg(), ocupadas:ocupadas()};
      case "pedir":        return pedir(p.cita || {});
      case "verCita":      return verCita(p.id, p.token);
      case "cancelarCita": return cancelarPorClienta(p.id, p.token);
      case "salon":        revisarClave(p.clave); return {ok:true, cfg:leerCfg(), citas:todas().map(sinToken)};
      case "estado":       revisarClave(p.clave); return cambiarPorSalon(p.id, p.estado, p.motivo);
      case "guardarCfg":   revisarClave(p.clave); return guardarCfg(p.cfg);
      default:             return {ok:true, mensaje:"Agenda de citas funcionando."};
    }
  } catch (err){
    return {ok:false, error:String(err && err.message || err)};
  }
}

function revisarClave(c){
  if (CLAVE_SALON === "CAMBIA-ESTA-CLAVE") throw new Error("El salón todavía no ha puesto su clave en el programa");
  if (String(c || "") !== CLAVE_SALON){ Utilities.sleep(1000); throw new Error("Clave incorrecta"); }
}

/* ---------- Hoja de cálculo ---------- */
function libro(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function hojaCitas(){
  let h = libro().getSheetByName(HOJA_CITAS);
  if (!h){
    h = libro().insertSheet(HOJA_CITAS);
    h.getRange(1, 1, 1000, COLS.length).setNumberFormat("@");
    h.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight("bold");
    h.setFrozenRows(1);
  }
  return h;
}
function texto(v){
  if (v instanceof Date) return Utilities.formatDate(v, ZONA, v.getFullYear() < 1900 ? "HH:mm" : "yyyy-MM-dd");
  return String(v);
}
function todas(){
  const h = hojaCitas(), n = h.getLastRow() - 1;
  if (n < 1) return [];
  return h.getRange(2, 1, n, COLS.length).getValues().map((f, i) => {
    const o = {fila: i + 2};
    COLS.forEach((k, j) => o[k] = texto(f[j]));
    o.min = Number(o.min);
    return o;
  });
}
function sinToken(c){ const o = Object.assign({}, c); delete o.token; delete o.fila; return o; }
function escribir(c){
  const h = hojaCitas(), fila = COLS.map(k => c[k] === undefined ? "" : String(c[k]));
  if (c.fila) h.getRange(c.fila, 1, 1, COLS.length).setValues([fila]);
  else { const r = h.getLastRow() + 1; h.getRange(r, 1, 1, COLS.length).setNumberFormat("@").setValues([fila]); }
}

/* ---------- Configuración ---------- */
function cfgInicial(){
  const h = {};
  for (let d = 0; d < 7; d++) h[d] = {abierto: d !== 0, desde:"08:00", hasta: d === 6 ? "14:00" : "18:00"};
  return {nombre:"Mi Salón", mensaje:"Escoge el servicio y la hora que te sirve.",
    servicios:[{id:"s1",icono:"💇",nombre:"Corte de cabello",min:30},{id:"s2",icono:"💈",nombre:"Barbería",min:30},
               {id:"s3",icono:"💅",nombre:"Uñas manos",min:60},{id:"s4",icono:"🎨",nombre:"Tintura",min:120}],
    horario:h, descanso:{activo:true, desde:"12:00", hasta:"13:00"},
    intervalo:30, simultaneas:1, diasAdelante:30, avisoMinHoras:2};
}
function leerCfg(){
  let h = libro().getSheetByName(HOJA_CONFIG);
  if (!h){ h = libro().insertSheet(HOJA_CONFIG); h.getRange("A1").setNumberFormat("@"); }
  const v = h.getRange("A1").getValue();
  if (!v){ const c = cfgInicial(); h.getRange("A1").setValue(JSON.stringify(c)); h.getRange("A2").setValue("No borres la celda A1: ahí está la configuración del salón."); return c; }
  return JSON.parse(v);
}
function guardarCfg(c){
  const hora = /^\d{2}:\d{2}$/;
  if (!c || !Array.isArray(c.servicios) || !c.servicios.length || c.servicios.length > 60) throw new Error("Revisa la lista de servicios");
  c.servicios = c.servicios.map(s => ({id:String(s.id).slice(0,20), icono:String(s.icono||"✨").slice(0,4), nombre:String(s.nombre||"").trim().slice(0,40), min:Math.min(600, Math.max(5, parseInt(s.min,10)||30))})).filter(s => s.nombre);
  for (let d = 0; d < 7; d++){ const x = c.horario && c.horario[d]; if (!x || !hora.test(x.desde) || !hora.test(x.hasta)) throw new Error("Revisa el horario"); }
  c.nombre = String(c.nombre || "Mi Salón").slice(0,50); c.mensaje = String(c.mensaje || "").slice(0,100);
  c.intervalo = [15,20,30,60].indexOf(+c.intervalo) >= 0 ? +c.intervalo : 30;
  c.simultaneas = Math.min(10, Math.max(1, +c.simultaneas || 1));
  c.diasAdelante = Math.min(120, Math.max(1, +c.diasAdelante || 30));
  c.avisoMinHoras = Math.min(48, Math.max(0, +c.avisoMinHoras || 0));
  leerCfg(); libro().getSheetByName(HOJA_CONFIG).getRange("A1").setValue(JSON.stringify(c));
  return {ok:true};
}

/* ---------- Horas libres (misma regla que la página) ---------- */
const aMin = s => { const p = s.split(":"); return +p[0] * 60 + +p[1]; };
const deMin = m => ("0" + Math.floor(m / 60)).slice(-2) + ":" + ("0" + m % 60).slice(-2);
function ahoraTxt(masHoras){ return Utilities.formatDate(new Date(Date.now() + (masHoras || 0) * 3600e3), ZONA, "yyyy-MM-dd HH:mm"); }
function horasLibres(c, lista, fecha, min){
  const p = fecha.split("-").map(Number), dia = new Date(p[0], p[1] - 1, p[2]).getDay(), h = c.horario[dia];
  if (!h || !h.abierto) return [];
  if (fecha > ahoraTxt(24 * c.diasAdelante).slice(0, 10)) return [];
  const limite = ahoraTxt(+c.avisoMinHoras || 0), paso = Math.max(5, +c.intervalo || 30), res = [];
  const ocup = lista.filter(x => x.fecha === fecha && (x.estado === "pendiente" || x.estado === "confirmada"));
  for (let t = aMin(h.desde); t + min <= aMin(h.hasta); t += paso){
    if (c.descanso && c.descanso.activo){ const a = aMin(c.descanso.desde), b = aMin(c.descanso.hasta); if (t < b && t + min > a) continue; }
    if (fecha + " " + deMin(t) < limite) continue;
    const cruces = ocup.filter(x => { const a = aMin(x.hora); return t < a + x.min && t + min > a; }).length;
    if (cruces >= Math.max(1, +c.simultaneas || 1)) continue;
    res.push(deMin(t));
  }
  return res;
}
function ocupadas(){
  const hoy = ahoraTxt().slice(0, 10);
  return todas().filter(x => x.fecha >= hoy && (x.estado === "pendiente" || x.estado === "confirmada"))
                .map(x => ({fecha:x.fecha, hora:x.hora, min:x.min, estado:"confirmada"}));
}

/* ---------- Citas ---------- */
function pedir(d){
  const c = leerCfg(), s = c.servicios.filter(x => x.id === d.servicioId)[0];
  if (!s) throw new Error("Ese servicio ya no está disponible");
  const nombre = String(d.nombre || "").trim().replace(/\s+/g, " ").slice(0, 60);
  const celular = String(d.celular || "").replace(/\D/g, "").replace(/^57(?=3\d{9}$)/, "");
  const correo = String(d.correo || "").trim().slice(0, 80), nota = String(d.nota || "").trim().slice(0, 200);
  if (nombre.length < 2) throw new Error("Escribe tu nombre");
  if (!/^3\d{9}$/.test(celular)) throw new Error("Escribe un celular de 10 dígitos que empiece por 3");
  if (correo && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(correo)) throw new Error("Ese correo no parece correcto");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.fecha || "") || !/^\d{2}:\d{2}$/.test(d.hora || "")) throw new Error("Escoge día y hora");
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const lista = todas();
    if (lista.filter(x => x.celular === celular && x.estado === "pendiente").length >= 3) throw new Error("Ya tienes 3 solicitudes esperando respuesta del salón");
    if (horasLibres(c, lista, d.fecha, s.min).indexOf(d.hora) < 0) throw new Error("Esa hora se acaba de ocupar. Escoge otra");
    const cita = {id:Utilities.getUuid().replace(/-/g, "").slice(0, 10), token:Utilities.getUuid().replace(/-/g, ""), estado:"pendiente",
      fecha:d.fecha, hora:d.hora, min:s.min, servicioId:s.id, servicio:s.nombre, icono:s.icono, nombre, celular, correo, nota,
      motivo:"", creada:ahoraTxt(), actualizada:""};
    escribir(cita);
    avisarSalon("Nueva solicitud: " + s.nombre + " " + fechaBonita(d.fecha) + " " + d.hora,
      nombre + " (" + celular + ") pidió " + s.nombre + " para el " + fechaBonita(d.fecha) + " a las " + d.hora + "." +
      (nota ? "\nNota: " + nota : "") + "\n\nAcéptala o recházala en tu panel: " + APP_URL + "#/salon");
    return {ok:true, id:cita.id, token:cita.token};
  } finally { lock.releaseLock(); }
}
function buscar(id, token){
  const x = todas().filter(c => c.id === String(id || ""))[0];
  if (!x || (token !== undefined && x.token !== String(token || ""))) throw new Error("No encontramos esa cita");
  return x;
}
function verCita(id, token){ const x = buscar(id, token); return {ok:true, cita:sinToken(x), cfg:leerCfg()}; }
function cancelarPorClienta(id, token){
  const x = buscar(id, token);
  if (x.estado !== "pendiente" && x.estado !== "confirmada") throw new Error("Esta cita ya no se puede cancelar");
  x.estado = "cancelada"; x.actualizada = ahoraTxt(); escribir(x);
  avisarSalon("Cita cancelada: " + x.servicio + " " + fechaBonita(x.fecha) + " " + x.hora, x.nombre + " canceló su cita. Esa hora quedó libre.");
  return {ok:true};
}
function cambiarPorSalon(id, estado, motivo){
  if (["confirmada", "rechazada", "cancelada"].indexOf(estado) < 0) throw new Error("Estado no válido");
  const x = buscar(id); x.estado = estado; x.motivo = String(motivo || "").slice(0, 120); x.actualizada = ahoraTxt(); escribir(x);
  if (x.correo){
    const nom = leerCfg().nombre, enlace = APP_URL + "#/cita/" + x.id + "/" + x.token;
    const txt = {confirmada:"¡Tu cita quedó CONFIRMADA! Te esperamos.", rechazada:"El salón no puede atenderte a esa hora." + (x.motivo ? " Motivo: " + x.motivo : "") + " Puedes pedir otra hora.", cancelada:"El salón canceló tu cita." + (x.motivo ? " Motivo: " + x.motivo : "")}[estado];
    try { MailApp.sendEmail({to:x.correo, name:nom, subject:nom + " – Tu cita: " + x.servicio + " " + fechaBonita(x.fecha) + " " + x.hora,
      body:"Hola " + x.nombre + ",\n\n" + txt + "\n\n" + x.servicio + " · " + fechaBonita(x.fecha) + " a las " + x.hora + "\n\nVer tu cita: " + enlace + "\n\n" + nom}); } catch(e){}
  }
  return {ok:true};
}
function avisarSalon(asunto, cuerpo){
  try { MailApp.sendEmail(Session.getEffectiveUser().getEmail(), asunto, cuerpo); } catch(e){}
}
function fechaBonita(f){
  const p = f.split("-").map(Number), d = new Date(p[0], p[1] - 1, p[2]);
  return ["dom","lun","mar","mié","jue","vie","sáb"][d.getDay()] + " " + p[2] + "/" + p[1];
}

/** Correr UNA vez desde el editor (botón Ejecutar) para dar permisos y crear las pestañas. */
function prepararHoja(){
  leerCfg(); hojaCitas();
  MailApp.getRemainingDailyQuota();
  Logger.log("Listo. Correos disponibles hoy: " + MailApp.getRemainingDailyQuota());
}
