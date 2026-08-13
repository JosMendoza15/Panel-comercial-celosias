import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  LayoutDashboard, ClipboardList, Users, Boxes, FileBarChart, BellRing,
  Plus, X, Pencil, Trash2, Search, Download, ChevronDown, ChevronRight,
  TrendingUp, Target, Wallet, Receipt, Package, Ruler, Users2, ShoppingCart,
  ArrowUpRight, ArrowDownRight, Clock, MapPin, Truck, Phone, Mail, Building2,
  CheckCircle2, AlertTriangle, CalendarClock, Sparkles, FileSpreadsheet, Printer,
  Layers, Palette, Wrench, Droplet, CreditCard, FileText, RefreshCw, Circle
} from "lucide-react";
import {
  ResponsiveContainer, LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, AreaChart, Area
} from "recharts";
import * as XLSX from "xlsx";

/* ============================== ADAPTADOR DE ALMACENAMIENTO ==============================
   Dentro de Claude usa window.storage. Publicada en tu propio dominio (fuera de Claude)
   usa localStorage del navegador. El resto del código llama siempre a "storage", sin
   preocuparse de dónde corre. */
const storage = (typeof window !== "undefined" && window.storage) ? window.storage : {
  async get(key) {
    const v = localStorage.getItem(key);
    return v !== null ? { key, value: v } : null;
  },
  async set(key, value) {
    localStorage.setItem(key, value);
    return { key, value };
  },
  async delete(key) {
    localStorage.removeItem(key);
    return { key, deleted: true };
  },
};

/* ============================== CONSTANTES ============================== */

const NAVY = "#0E2A47";
const NAVY_LIGHT = "#1B4A78";
const NAVY_SOFT = "#274F79";
const PAPER = "#F4F6F8";
const LINE = "#E2E7EC";
const INK = "#152238";
const MUTED = "#5C6B7F";
const GOOD = "#1F9E77";
const WARN = "#C88A2C";
const BAD = "#C2483F";
const ACCENT = "#3E7CB8";

const TIERS = [
  { min: 0, pct: 0 },
  { min: 250000, pct: 2.0 },
  { min: 350000, pct: 2.25 },
  { min: 450000, pct: 2.5 },
  { min: 550000, pct: 3.0 },
  { min: 650000, pct: 3.25 },
  { min: 700000, pct: 3.5 },
];
const META_MENSUAL = 250000;

const ETAPAS = [
  "Prospecto", "Cotización enviada", "Negociación", "Pedido confirmado",
  "Producción", "Embarcado", "Entregado", "Cobrado",
];
const ETAPAS_TEMPRANAS = ["Prospecto", "Cotización enviada", "Negociación"];
const ETAPA_COLOR = {
  "Prospecto": "#8C99AB",
  "Cotización enviada": "#3E7CB8",
  "Negociación": "#C88A2C",
  "Pedido confirmado": "#3E63B8",
  "Producción": "#7A4FC2",
  "Embarcado": "#2C8FAE",
  "Entregado": "#2C8F5B",
  "Cobrado": "#1F9E77",
};

const DEFAULT_MODELOS = [
  "Celosía Bloque Recto", "Celosía Rombo", "Celosía Ola", "Celosía Panal",
  "Celosía Circular", "Celosía Cuadrícula", "Celosía Colonial", "Celosía Minimal",
];
const DEFAULT_COLORES = ["Gris natural", "Blanco", "Arena", "Negro", "Terracota", "Gris con óxido negro", "Terracota con óxido rojo"];
const DEFAULT_ACABADOS = ["Pulido", "Rústico", "Liso", "Martelinado"];
const ESTADOS_MX = [
  "Aguascalientes","Baja California","Baja California Sur","Campeche","Chiapas","Chihuahua",
  "Ciudad de México","Coahuila","Colima","Durango","Estado de México","Guanajuato","Guerrero",
  "Hidalgo","Jalisco","Michoacán","Morelos","Nayarit","Nuevo León","Oaxaca","Puebla","Querétaro",
  "Quintana Roo","San Luis Potosí","Sinaloa","Sonora","Tabasco","Tamaulipas","Tlaxcala",
  "Veracruz","Yucatán","Zacatecas",
];
const DEFAULT_FLETERAS = ["Estafeta", "Paquetexpress", "Tres Guerras", "Flete propio", "Transporte del cliente", "Otra"];
const DEFAULT_FORMAS_PAGO = ["Transferencia", "Efectivo", "Cheque", "Tarjeta", "Depósito"];

const CATALOG_DEFS = [
  { key: "modelos", label: "Modelos", icon: "Boxes", field: "modelo" },
  { key: "colores", label: "Colores / Paleta", icon: "Palette", field: "color" },
  { key: "acabados", label: "Acabados", icon: "Wrench", field: "acabado" },
  { key: "fleteras", label: "Fleteras", icon: "Truck", field: "fletera" },
  { key: "formasPago", label: "Formas de pago", icon: "CreditCard", field: "formaPago" },
];

const MESES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

/* ============================== UTILIDADES ============================== */

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

const fmtMoney = (n) =>
  (isFinite(n) ? n : 0).toLocaleString("es-MX", { style: "currency", currency: "MXN", minimumFractionDigits: 2, maximumFractionDigits: 2 });

const fmtNum = (n, d = 0) => (isFinite(n) ? n : 0).toLocaleString("es-MX", { maximumFractionDigits: d });

const fmtDate = (iso) => {
  if (!iso) return "—";
  const d = new Date(iso.slice(0, 10) + "T00:00:00");
  if (isNaN(d)) return iso;
  return d.toLocaleDateString("es-MX", { day: "2-digit", month: "short", year: "numeric" });
};

const fmtDateTime = (iso) => {
  if (!iso) return "—";
  const d = new Date(iso.length > 10 ? iso : iso + "T00:00");
  if (isNaN(d)) return iso;
  const soloFecha = iso.length <= 10;
  return d.toLocaleString("es-MX", soloFecha
    ? { day: "2-digit", month: "short", year: "numeric" }
    : { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" });
};

const todayISO = () => new Date().toISOString().slice(0, 10);

const nowLocalISO = () => {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
};

const addDaysLocalISO = (days) => {
  const d = new Date();
  d.setDate(d.getDate() + days);
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
};

const monthKeyOf = (iso) => (iso ? iso.slice(0, 7) : "");

const calcSubtotalIVA = (total) => {
  const t = Number(total) || 0;
  const subtotal = t / 1.16;
  const iva = t - subtotal;
  return { subtotal, iva };
};

const getTier = (subtotal) => {
  let current = TIERS[0];
  for (const t of TIERS) if (subtotal >= t.min) current = t;
  return current;
};

const getNextTier = (subtotal) => {
  const next = TIERS.find((t) => t.min > subtotal);
  return next || null;
};

const daysInMonth = (year, monthIdx) => new Date(year, monthIdx + 1, 0).getDate();

const round2 = (n) => Math.round((Number(n) || 0) * 100) / 100;

// Imprime con un título de documento limpio (sin nombres técnicos de la app).
// El navegador puede agregar URL en encabezado/pie: recuerda a la persona
// desactivar "Encabezados y pies de página" en el diálogo de impresión.
const printClean = (title) => {
  const prevTitle = document.title;
  document.title = title || "Reporte";
  window.print();
  setTimeout(() => { document.title = prevTitle; }, 800);
};

const TIPOS_MOVIMIENTO = ["Anticipo", "Pago parcial", "Finiquito"];

// Compatibilidad: pedidos capturados antes de este cambio no tienen "pagos",
// así que se reconstruye un único movimiento a partir de los campos viejos.
const getPagosPedido = (v) => {
  if (Array.isArray(v.pagos) && v.pagos.length) return v.pagos;
  if (v.etapa === "Cobrado" && Number(v.totalCobrado) > 0) {
    const { subtotal } = calcSubtotalIVA(v.totalCobrado);
    return [{ id: `legacy-${v.id}`, fecha: v.fechaRealCobro || v.fecha, monto: round2(subtotal), montoRecibido: Number(v.totalCobrado), conIVA: true, tipo: "Finiquito" }];
  }
  return [];
};

const getSubtotalPedido = (v) => {
  if (v.subtotal !== undefined && v.subtotal !== null && v.subtotal !== "") return Number(v.subtotal) || 0;
  if (Number(v.totalCobrado) > 0) return calcSubtotalIVA(v.totalCobrado).subtotal; // compatibilidad
  return 0;
};

const getTotalPagado = (v) => getPagosPedido(v).reduce((s, p) => s + (Number(p.monto) || 0), 0);
const getSaldoPendiente = (v) => round2(getSubtotalPedido(v) - getTotalPagado(v));
const getLiquidado = (v) => getPagosPedido(v).length > 0 && getSaldoPendiente(v) <= 0.5;

/* ============================== DATOS DE EJEMPLO ============================== */

function seedData() {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth();
  const rows = [];
  const clientes = [
    { cliente: "Arturo Villanueva", empresa: "Constructora Villanueva", telefono: "9991234567", correo: "arturo@cvillanueva.mx" },
    { cliente: "Mariana Cetz", empresa: "Grupo Cetz Inmobiliaria", telefono: "9992345678", correo: "mariana@grupocetz.mx" },
    { cliente: "Ramón Uc", empresa: "Materiales Uc", telefono: "9993456789", correo: "ramon@materialesuc.mx" },
    { cliente: "Fernanda Aguilar", empresa: "Aguilar Diseño y Obra", telefono: "9994567890", correo: "fer@aguilardiseno.mx" },
    { cliente: "Luis Pat", empresa: "Pat Construcciones", telefono: "9995678901", correo: "luis@patconstrucciones.mx" },
  ];
  const combos = [
    { d: 2, cli: 0, etapa: "Cobrado", total: 98600, piezas: 420, m2: 84, estado: "Yucatán", ciudad: "Mérida", envio: "Local" },
    { d: 5, cli: 1, etapa: "Cobrado", total: 145000, piezas: 610, m2: 122, estado: "Quintana Roo", ciudad: "Cancún", envio: "Nacional" },
    { d: 8, cli: 2, etapa: "Cobrado", total: 76300, piezas: 300, m2: 60, estado: "Yucatán", ciudad: "Progreso", envio: "Local" },
    { d: 11, cli: 3, etapa: "Producción", total: 132000, piezas: 540, m2: 108, estado: "Campeche", ciudad: "Campeche", envio: "Nacional" },
    { d: 13, cli: 0, etapa: "Cobrado", total: 61000, piezas: 250, m2: 50, estado: "Yucatán", ciudad: "Mérida", envio: "Local" },
    { d: 16, cli: 4, etapa: "Embarcado", total: 88500, piezas: 360, m2: 72, estado: "Tabasco", ciudad: "Villahermosa", envio: "Nacional" },
    { d: 18, cli: 1, etapa: "Cobrado", total: 54200, piezas: 220, m2: 44, estado: "Quintana Roo", ciudad: "Playa del Carmen", envio: "Nacional" },
    { d: 20, cli: 2, etapa: "Negociación", total: 47000, piezas: 190, m2: 38, estado: "Yucatán", ciudad: "Umán", envio: "Local" },
    { d: 22, cli: 3, etapa: "Cobrado", total: 39800, piezas: 160, m2: 32, estado: "Yucatán", ciudad: "Mérida", envio: "Local" },
    { d: 24, cli: 4, etapa: "Cotización enviada", total: 210000, piezas: 860, m2: 172, estado: "Quintana Roo", ciudad: "Tulum", envio: "Nacional" },
  ];
  combos.forEach((c, i) => {
    const cli = clientes[c.cli];
    const fecha = new Date(y, m, c.d).toISOString().slice(0, 10);
    rows.push({
      id: uid(),
      fecha,
      numeroPedido: `PED-${1000 + i}`,
      cliente: cli.cliente,
      empresa: cli.empresa,
      telefono: cli.telefono,
      correo: cli.correo,
      modelo: DEFAULT_MODELOS[i % DEFAULT_MODELOS.length],
      cantidadPiezas: c.piezas,
      metrosCuadrados: c.m2,
      color: DEFAULT_COLORES[i % DEFAULT_COLORES.length],
      acabado: DEFAULT_ACABADOS[i % DEFAULT_ACABADOS.length],
      estadoDestino: c.estado,
      ciudadDestino: c.ciudad,
      tipoEnvio: c.envio,
      fletera: DEFAULT_FLETERAS[i % DEFAULT_FLETERAS.length],
      totalCobrado: c.total,
      formaPago: DEFAULT_FORMAS_PAGO[i % DEFAULT_FORMAS_PAGO.length],
      anticipo: Math.round(c.total * 0.3),
      saldoPendiente: c.etapa === "Cobrado" ? 0 : Math.round(c.total * 0.7),
      fechaEstimadaCobro: fecha,
      fechaRealCobro: c.etapa === "Cobrado" ? fecha : "",
      observaciones: "",
      etapa: c.etapa,
      creadoEn: fecha,
    });
  });
  return rows;
}

/* ============================== SINCRONIZACIÓN CON SUPABASE (nube) ==============================
   Capa adicional: si hay sesión iniciada, los datos se leen/guardan en Supabase
   (accesible desde cualquier dispositivo con tu usuario). Si no hay sesión,
   todo sigue funcionando igual que antes con el almacenamiento local de Claude.
   ================================================================================== */

const SUPABASE_URL = "https://jljnloekrnuqeytdlqsp.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_O9r3OEOanO0xQ0Ug2vcq8A_RmU5NTiW";
const SESSION_CACHE_KEY = "supabase-session-cache-v1";

let currentSession = null; // { access_token, refresh_token, email } | null

async function cacheSession(session) {
  try {
    if (session) await storage.set(SESSION_CACHE_KEY, JSON.stringify(session), false);
    else await storage.delete(SESSION_CACHE_KEY, false);
  } catch (e) { /* no-op */ }
}

async function loadCachedSession() {
  try {
    const res = await storage.get(SESSION_CACHE_KEY, false);
    if (res && res.value) return JSON.parse(res.value);
    return null;
  } catch (e) { return null; }
}

// El componente App conecta esto a un estado visible, para que los errores de sync ya no sean silenciosos.
let onSyncError = null;
let onSyncOk = null;
function reportSyncError(msg) { if (onSyncError) onSyncError(msg); }
function reportSyncOk() { if (onSyncOk) onSyncOk(); }

async function supaLogin(email, password) {
  const res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=password`, {
    method: "POST",
    headers: { "Content-Type": "application/json", apikey: SUPABASE_ANON_KEY },
    body: JSON.stringify({ email, password }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error_description || data.msg || "No se pudo iniciar sesión. Revisa tu correo y contraseña.");
  return { access_token: data.access_token, refresh_token: data.refresh_token, email };
}

async function supaRefreshToken() {
  if (!currentSession || !currentSession.refresh_token) return false;
  try {
    const res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=refresh_token`, {
      method: "POST",
      headers: { "Content-Type": "application/json", apikey: SUPABASE_ANON_KEY },
      body: JSON.stringify({ refresh_token: currentSession.refresh_token }),
    });
    const data = await res.json();
    if (!res.ok || !data.access_token) return false;
    currentSession = { access_token: data.access_token, refresh_token: data.refresh_token, email: currentSession.email };
    await cacheSession(currentSession);
    return true;
  } catch (e) {
    return false;
  }
}

function supaHeaders(extra = {}) {
  return {
    apikey: SUPABASE_ANON_KEY,
    Authorization: `Bearer ${currentSession ? currentSession.access_token : SUPABASE_ANON_KEY}`,
    "Content-Type": "application/json",
    ...extra,
  };
}

async function supaFetch(url, options = {}, isRetry = false) {
  const res = await fetch(url, { ...options, headers: supaHeaders(options.headers) });
  if (res.status === 401 && currentSession && !isRetry) {
    const refreshed = await supaRefreshToken();
    if (refreshed) return supaFetch(url, options, true);
  }
  return res;
}

async function extractError(res, fallback) {
  try {
    const data = await res.json();
    return data.message || data.error_description || data.hint || fallback;
  } catch (e) {
    return fallback;
  }
}

async function supaSelect(table) {
  const res = await supaFetch(`${SUPABASE_URL}/rest/v1/${table}?select=*`);
  if (!res.ok) {
    const msg = await extractError(res, `No se pudo leer "${table}" (${res.status})`);
    reportSyncError(msg);
    throw new Error(msg);
  }
  reportSyncOk();
  return res.json();
}

async function supaUpsert(table, rows) {
  const res = await supaFetch(`${SUPABASE_URL}/rest/v1/${table}`, {
    method: "POST",
    headers: { Prefer: "resolution=merge-duplicates,return=minimal" },
    body: JSON.stringify(Array.isArray(rows) ? rows : [rows]),
  });
  if (!res.ok) {
    const msg = await extractError(res, `No se pudo guardar en "${table}" (${res.status})`);
    reportSyncError(msg);
    throw new Error(msg);
  }
  reportSyncOk();
}

async function supaDeleteRow(table, id) {
  const res = await supaFetch(`${SUPABASE_URL}/rest/v1/${table}?id=eq.${encodeURIComponent(id)}`, { method: "DELETE" });
  if (!res.ok) {
    const msg = await extractError(res, `No se pudo borrar en "${table}" (${res.status})`);
    reportSyncError(msg);
    throw new Error(msg);
  }
  reportSyncOk();
}

async function supaReplaceTable(table, rows) {
  // Primero se sube/actualiza lo nuevo (si esto falla, no se pierde nada de lo que ya había).
  if (rows.length) await supaUpsert(table, rows);
  // Solo después se borran en Supabase las filas que ya no existen en "rows" (p. ej. las que se eliminaron en la app).
  const idsActuales = rows.map((r) => r.id);
  const existentes = await supaSelect(table).catch(() => []);
  const aBorrar = existentes.filter((r) => !idsActuales.includes(r.id));
  for (const r of aBorrar) {
    await supaDeleteRow(table, r.id).catch((e) => console.error(`No se pudo borrar fila obsoleta en ${table}`, e));
  }
}

function ventaToRow(v) {
  return {
    id: v.id, fecha: v.fecha || null, numero_pedido: v.numeroPedido, cliente: v.cliente, empresa: v.empresa,
    telefono: v.telefono, correo: v.correo, modelo: v.modelo, cantidad_piezas: v.cantidadPiezas,
    metros_cuadrados: v.metrosCuadrados, color: v.color, acabado: v.acabado, estado_destino: v.estadoDestino,
    ciudad_destino: v.ciudadDestino, tipo_envio: v.tipoEnvio, fletera: v.fletera, total_cobrado: v.totalCobrado,
    forma_pago: v.formaPago, anticipo: v.anticipo, saldo_pendiente: v.saldoPendiente,
    fecha_estimada_cobro: v.fechaEstimadaCobro || null, fecha_real_cobro: v.fechaRealCobro || null,
    observaciones: v.observaciones, etapa: v.etapa, proximo_seguimiento: v.proximoSeguimiento || null,
    nota_seguimiento: v.notaSeguimiento,
    subtotal: v.subtotal === "" || v.subtotal === undefined ? null : v.subtotal,
    con_iva: !!v.conIVA, anticipo_pct: v.anticipoPct === "" || v.anticipoPct === undefined ? null : v.anticipoPct,
    pagos: JSON.stringify(Array.isArray(v.pagos) ? v.pagos : []),
  };
}
function rowToVenta(r) {
  return {
    id: r.id, fecha: r.fecha, numeroPedido: r.numero_pedido, cliente: r.cliente, empresa: r.empresa,
    telefono: r.telefono, correo: r.correo, modelo: r.modelo, cantidadPiezas: r.cantidad_piezas,
    metrosCuadrados: r.metros_cuadrados, color: r.color, acabado: r.acabado, estadoDestino: r.estado_destino,
    ciudadDestino: r.ciudad_destino, tipoEnvio: r.tipo_envio, fletera: r.fletera, totalCobrado: r.total_cobrado,
    formaPago: r.forma_pago, anticipo: r.anticipo, saldoPendiente: r.saldo_pendiente,
    fechaEstimadaCobro: r.fecha_estimada_cobro, fechaRealCobro: r.fecha_real_cobro,
    observaciones: r.observaciones, etapa: r.etapa, proximoSeguimiento: r.proximo_seguimiento,
    notaSeguimiento: r.nota_seguimiento,
    subtotal: r.subtotal, conIVA: !!r.con_iva, anticipoPct: r.anticipo_pct,
    pagos: typeof r.pagos === "string" ? (JSON.parse(r.pagos || "[]")) : (Array.isArray(r.pagos) ? r.pagos : []),
  };
}

function oportunidadToRow(o) {
  return {
    id: o.id, cliente: o.cliente, empresa: o.empresa, ciudad: o.ciudad, telefono: o.telefono,
    valor_estimado: o.valorEstimado, fecha_ultimo_contacto: o.fechaUltimoContacto || null,
    proxima_accion: o.proximaAccion, fecha_proxima_accion: o.fechaProximaAccion || null,
    prioridad: o.prioridad, notas: o.notas, etapa: o.etapa, seguimientos: o.seguimientos || [],
    modelo: o.modelo || "", motivo_perdida: o.motivoPerdida || "",
  };
}
function rowToOportunidad(r) {
  return {
    id: r.id, cliente: r.cliente, empresa: r.empresa, ciudad: r.ciudad, telefono: r.telefono,
    valorEstimado: r.valor_estimado, fechaUltimoContacto: r.fecha_ultimo_contacto,
    proximaAccion: r.proxima_accion, fechaProximaAccion: r.fecha_proxima_accion,
    prioridad: r.prioridad, notas: r.notas, etapa: r.etapa, creadoEn: r.creado_en,
    seguimientos: r.seguimientos || [], modelo: r.modelo || "", motivoPerdida: r.motivo_perdida || "",
  };
}

/* ============================== ALMACENAMIENTO ============================== */

const STORAGE_KEY = "ventas-celosias-v1";
const CATALOG_KEY = "catalogos-celosias-v1";

const DEFAULT_CATALOGOS = {
  modelos: DEFAULT_MODELOS,
  colores: DEFAULT_COLORES,
  acabados: DEFAULT_ACABADOS,
  fleteras: DEFAULT_FLETERAS,
  formasPago: DEFAULT_FORMAS_PAGO,
  empresaInfo: { nombre: "", telefono: "", correo: "" },
};

async function loadFromStorage() {
  if (currentSession) {
    try {
      const rows = await supaSelect("ventas");
      if (rows.length > 0) {
        const ventas = rows.map(rowToVenta);
        await storage.set(STORAGE_KEY, JSON.stringify(ventas), false).catch(() => {});
        return ventas;
      }
      const localRes = await storage.get(STORAGE_KEY, false).catch(() => null);
      const local = localRes && localRes.value ? JSON.parse(localRes.value) : null;
      if (Array.isArray(local) && local.length) {
        await supaReplaceTable("ventas", local.map(ventaToRow)).catch((e) => console.error("No se pudo subir a Supabase", e));
        return local;
      }
      return [];
    } catch (e) {
      console.error("Fallo Supabase, usando copia local", e);
    }
  }
  try {
    const res = await storage.get(STORAGE_KEY, false);
    if (res && res.value) return JSON.parse(res.value);
    return null;
  } catch (e) {
    return null;
  }
}

async function saveToStorage(ventas) {
  try {
    await storage.set(STORAGE_KEY, JSON.stringify(ventas), false);
  } catch (e) {
    console.error("No se pudo guardar localmente", e);
  }
  if (currentSession) {
    try {
      await supaReplaceTable("ventas", ventas.map(ventaToRow));
    } catch (e) {
      console.error("No se pudo sincronizar ventas con Supabase", e);
      return false;
    }
  }
  return true;
}

async function loadCatalogos() {
  if (currentSession) {
    try {
      const res = await supaFetch(`${SUPABASE_URL}/rest/v1/app_config?key=eq.catalogos&select=value`);
      if (res.ok) {
        const rows = await res.json();
        if (rows.length) {
          await storage.set(CATALOG_KEY, JSON.stringify(rows[0].value), false).catch(() => {});
          return rows[0].value;
        }
        const localRes = await storage.get(CATALOG_KEY, false).catch(() => null);
        const local = localRes && localRes.value ? JSON.parse(localRes.value) : null;
        if (local) {
          await supaUpsert("app_config", { key: "catalogos", value: local }).catch((e) => console.error("No se pudo subir a Supabase", e));
          return local;
        }
      }
    } catch (e) {
      console.error("Fallo Supabase, usando copia local", e);
    }
  }
  try {
    const res = await storage.get(CATALOG_KEY, false);
    if (res && res.value) return JSON.parse(res.value);
    return null;
  } catch (e) {
    return null;
  }
}

async function saveCatalogos(cat) {
  try {
    await storage.set(CATALOG_KEY, JSON.stringify(cat), false);
  } catch (e) {
    console.error("No se pudo guardar localmente", e);
  }
  if (currentSession) {
    try {
      await supaUpsert("app_config", { key: "catalogos", value: cat });
    } catch (e) {
      console.error("No se pudo sincronizar catálogos con Supabase", e);
      return false;
    }
  }
  return true;
}

/* ============================== ICONO CELOSÍA (firma visual) ============================== */

function CelosiaPattern({ id = "celosia" }) {
  return (
    <defs>
      <pattern id={id} width="28" height="28" patternUnits="userSpaceOnUse">
        <rect width="28" height="28" fill="none" />
        <rect x="4" y="4" width="20" height="20" rx="3" fill="none" stroke="rgba(255,255,255,0.14)" strokeWidth="2" />
      </pattern>
    </defs>
  );
}

function CelosiaBand({ height = 64 }) {
  return (
    <svg width="100%" height={height} style={{ position: "absolute", inset: 0, opacity: 0.55 }} preserveAspectRatio="none">
      <CelosiaPattern id="celosia-band" />
      <rect width="100%" height="100%" fill="url(#celosia-band)" />
    </svg>
  );
}

/* ============================== COMPONENTES BASE ============================== */

function Card({ children, style, className = "" }) {
  return (
    <div
      className={className}
      style={{
        background: "#FFFFFF",
        border: `1px solid ${LINE}`,
        borderRadius: 14,
        boxShadow: "0 1px 2px rgba(14,42,71,0.04)",
        ...style,
      }}
    >
      {children}
    </div>
  );
}

function StatCard({ icon: Icon, label, value, sub, accent = NAVY, trend }) {
  return (
    <Card style={{ padding: "16px 18px", position: "relative", overflow: "hidden" }}>
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between" }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 6, minWidth: 0 }}>
          <span style={{ fontSize: 12, fontWeight: 600, color: MUTED, letterSpacing: 0.3, textTransform: "uppercase" }}>{label}</span>
          <span style={{ fontSize: 22, fontWeight: 700, color: INK, lineHeight: 1.1, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{value}</span>
          {sub && <span style={{ fontSize: 12.5, color: MUTED }}>{sub}</span>}
        </div>
        <div style={{ width: 36, height: 36, borderRadius: 10, background: `${accent}14`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
          <Icon size={18} color={accent} />
        </div>
      </div>
      {trend != null && (
        <div style={{ marginTop: 10, display: "flex", alignItems: "center", gap: 4, fontSize: 12.5, color: trend >= 0 ? GOOD : BAD, fontWeight: 600 }}>
          {trend >= 0 ? <ArrowUpRight size={14} /> : <ArrowDownRight size={14} />}
          {Math.abs(trend).toFixed(1)}% vs mes anterior
        </div>
      )}
    </Card>
  );
}

function Pill({ children, color = NAVY, bg, style }) {
  return (
    <span
      style={{
        display: "inline-flex", alignItems: "center", gap: 5, padding: "3px 10px",
        borderRadius: 999, fontSize: 11.5, fontWeight: 700, color,
        background: bg || `${color}16`, whiteSpace: "nowrap",
        ...style,
      }}
    >
      {children}
    </span>
  );
}

function Button({ children, onClick, variant = "primary", size = "md", icon: Icon, style, disabled, type = "button" }) {
  const base = {
    display: "inline-flex", alignItems: "center", gap: 7, justifyContent: "center",
    borderRadius: 10, fontWeight: 600, cursor: disabled ? "not-allowed" : "pointer",
    border: "1px solid transparent", transition: "filter 0.15s ease, transform 0.05s ease",
    fontSize: size === "sm" ? 12.5 : 13.5, padding: size === "sm" ? "6px 10px" : "9px 14px",
    opacity: disabled ? 0.55 : 1,
  };
  const variants = {
    primary: { background: NAVY, color: "#fff" },
    secondary: { background: "#fff", color: NAVY, border: `1px solid ${LINE}` },
    ghost: { background: "transparent", color: MUTED },
    danger: { background: "#fff", color: BAD, border: `1px solid #F0D4D2` },
    subtle: { background: PAPER, color: INK },
  };
  return (
    <button
      type={type}
      disabled={disabled}
      onClick={onClick}
      onMouseDown={(e) => { if (!disabled) e.currentTarget.style.transform = "scale(0.98)"; }}
      onMouseUp={(e) => (e.currentTarget.style.transform = "scale(1)")}
      style={{ ...base, ...variants[variant], ...style }}
    >
      {Icon && <Icon size={size === "sm" ? 14 : 15} />}
      {children}
    </button>
  );
}

function Field({ label, children, span }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 5, gridColumn: span ? `span ${span}` : undefined }}>
      <span style={{ fontSize: 12, fontWeight: 600, color: MUTED }}>{label}</span>
      {children}
    </label>
  );
}

const inputStyle = {
  border: `1px solid ${LINE}`, borderRadius: 9, padding: "8px 10px", fontSize: 13.5,
  color: INK, background: "#fff", outline: "none", width: "100%", fontFamily: "inherit",
};

function TextInput(props) {
  return <input {...props} style={{ ...inputStyle, ...(props.style || {}) }} />;
}
function Select({ children, ...props }) {
  return <select {...props} style={{ ...inputStyle, ...(props.style || {}) }}>{children}</select>;
}
function TextArea(props) {
  return <textarea {...props} style={{ ...inputStyle, resize: "vertical", minHeight: 60, ...(props.style || {}) }} />;
}

function EmptyState({ icon: Icon, title, sub }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "48px 20px", color: MUTED, gap: 8 }}>
      <Icon size={30} color={LINE === "#E2E7EC" ? "#C7D0DA" : MUTED} />
      <div style={{ fontWeight: 700, color: INK, fontSize: 14 }}>{title}</div>
      {sub && <div style={{ fontSize: 12.5, textAlign: "center", maxWidth: 320 }}>{sub}</div>}
    </div>
  );
}

function CatalogSection({ title, icon: Icon, items, usageCount, onAdd, onRemove }) {
  const [newItem, setNewItem] = useState("");
  const submit = () => {
    if (newItem.trim()) {
      onAdd(newItem.trim());
      setNewItem("");
    }
  };
  return (
    <Card style={{ padding: 16 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
        <div style={{ width: 30, height: 30, borderRadius: 8, background: `${NAVY}14`, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <Icon size={15} color={NAVY} />
        </div>
        <div style={{ fontWeight: 700, fontSize: 14 }}>{title}</div>
        <span style={{ marginLeft: "auto", fontSize: 11, color: MUTED }}>{items.length} elemento(s)</span>
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 12, minHeight: 28 }}>
        {items.map((item, i) => {
          const uso = usageCount ? usageCount(item) : 0;
          return (
            <span key={item + i} style={{
              display: "inline-flex", alignItems: "center", gap: 6, padding: "5px 6px 5px 10px",
              borderRadius: 999, background: PAPER, border: `1px solid ${LINE}`, fontSize: 12.5, color: INK,
            }}>
              {item}
              {uso > 0 && <span style={{ fontSize: 10.5, color: MUTED }}>({uso})</span>}
              <button
                onClick={() => onRemove(i)}
                title={uso > 0 ? "Ya se usa en pedidos existentes; se puede eliminar del catálogo de todos modos" : "Eliminar"}
                style={{ border: "none", background: "transparent", cursor: "pointer", color: MUTED, display: "flex", padding: 2 }}
              >
                <X size={12} />
              </button>
            </span>
          );
        })}
        {!items.length && <span style={{ fontSize: 12.5, color: MUTED }}>Sin elementos todavía.</span>}
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <TextInput
          placeholder={`Agregar a ${title.toLowerCase()}…`}
          value={newItem}
          onChange={(e) => setNewItem(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
        />
        <Button size="sm" icon={Plus} onClick={submit}>Agregar</Button>
      </div>
    </Card>
  );
}

function KanbanCard({ v, onEdit, onDragStart }) {
  const total = getSubtotalPedido(v);
  const vencido = v.proximoSeguimiento && v.proximoSeguimiento <= nowLocalISO();
  return (
    <div
      draggable
      onDragStart={(e) => { e.dataTransfer.effectAllowed = "move"; onDragStart(v.id); }}
      onClick={() => onEdit(v)}
      style={{
        background: "#fff", border: `1px solid ${LINE}`, borderLeft: `3px solid ${ETAPA_COLOR[v.etapa]}`,
        borderRadius: 9, padding: "9px 10px", cursor: "grab", boxShadow: "0 1px 2px rgba(14,42,71,0.05)",
      }}
    >
      <div style={{ fontWeight: 700, fontSize: 12.5, color: INK }}>{v.cliente || "Sin nombre"}</div>
      {v.empresa && <div style={{ fontSize: 11, color: MUTED, marginTop: 1 }}>{v.empresa}</div>}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 6, fontSize: 11, color: MUTED }}>
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 110 }}>{v.modelo}</span>
        <span style={{ fontWeight: 700, color: total > 0 ? INK : MUTED }}>{total > 0 ? fmtMoney(total) : "Sin monto"}</span>
      </div>
      {v.proximoSeguimiento && (
        <div style={{ display: "flex", alignItems: "center", gap: 4, marginTop: 6, fontSize: 10.5, color: vencido ? BAD : MUTED, fontWeight: vencido ? 700 : 500 }}>
          <CalendarClock size={11} /> {fmtDateTime(v.proximoSeguimiento)}{v.notaSeguimiento ? ` · ${v.notaSeguimiento}` : ""}
        </div>
      )}
    </div>
  );
}

function KanbanBoard({ ventas, onStageChange, onEdit }) {
  const [dragId, setDragId] = useState(null);
  const [overCol, setOverCol] = useState(null);
  return (
    <div style={{ display: "flex", gap: 10, overflowX: "auto", paddingBottom: 6 }}>
      {ETAPAS.map((etapa) => {
        const items = ventas.filter((v) => v.etapa === etapa);
        return (
          <div
            key={etapa}
            onDragOver={(e) => { e.preventDefault(); setOverCol(etapa); }}
            onDragLeave={() => setOverCol((c) => (c === etapa ? null : c))}
            onDrop={(e) => {
              e.preventDefault();
              if (dragId) onStageChange(dragId, etapa);
              setDragId(null); setOverCol(null);
            }}
            style={{
              minWidth: 230, width: 230, flexShrink: 0, background: overCol === etapa ? `${ETAPA_COLOR[etapa]}0F` : PAPER,
              borderRadius: 12, border: `1px solid ${overCol === etapa ? ETAPA_COLOR[etapa] : LINE}`,
              display: "flex", flexDirection: "column", maxHeight: "68vh",
            }}
          >
            <div style={{ padding: "10px 12px", borderBottom: `2px solid ${ETAPA_COLOR[etapa]}`, display: "flex", alignItems: "center", gap: 6 }}>
              <span style={{ width: 8, height: 8, borderRadius: 999, background: ETAPA_COLOR[etapa], flexShrink: 0 }} />
              <span style={{ fontWeight: 700, fontSize: 12, color: INK }}>{etapa}</span>
              <span style={{ marginLeft: "auto", fontSize: 11, color: MUTED, background: "#fff", borderRadius: 999, padding: "1px 7px", border: `1px solid ${LINE}` }}>{items.length}</span>
            </div>
            <div style={{ padding: 8, display: "flex", flexDirection: "column", gap: 8, overflowY: "auto" }}>
              {items.map((v) => <KanbanCard key={v.id} v={v} onEdit={onEdit} onDragStart={setDragId} />)}
              {!items.length && <div style={{ fontSize: 11, color: MUTED, textAlign: "center", padding: "16px 0" }}>Sin pedidos aquí</div>}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ============================== INICIO DE SESIÓN (Supabase) ============================== */

function LoginScreen({ onLogin, onOfflineMode }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loadingLogin, setLoadingLogin] = useState(false);
  const [error, setError] = useState("");

  const submit = async () => {
    if (!email.trim() || !password) return;
    setLoadingLogin(true);
    setError("");
    try {
      await onLogin(email.trim(), password);
    } catch (e) {
      setError(e.message || "No se pudo iniciar sesión.");
    }
    setLoadingLogin(false);
  };

  return (
    <div style={{
      minHeight: 480, display: "flex", alignItems: "center", justifyContent: "center", padding: 24,
      background: PAPER, borderRadius: 16, border: `1px solid ${LINE}`,
      fontFamily: "'Segoe UI', 'Inter', system-ui, -apple-system, sans-serif",
    }}>
      <Card style={{ padding: 28, width: "100%", maxWidth: 360 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
          <div style={{ width: 34, height: 34, borderRadius: 9, background: NAVY, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <svg width="17" height="17" viewBox="0 0 20 20" fill="none">
              <rect x="1" y="1" width="8" height="8" rx="1.5" stroke="#fff" strokeWidth="1.6" />
              <rect x="11" y="1" width="8" height="8" rx="1.5" stroke="#fff" strokeWidth="1.6" />
              <rect x="1" y="11" width="8" height="8" rx="1.5" stroke="#fff" strokeWidth="1.6" />
              <rect x="11" y="11" width="8" height="8" rx="1.5" stroke="#fff" strokeWidth="1.6" />
            </svg>
          </div>
          <div style={{ fontWeight: 800, fontSize: 15.5, color: INK }}>Panel Comercial · Celosías</div>
        </div>
        <div style={{ fontSize: 12.5, color: MUTED, marginBottom: 18 }}>
          Inicia sesión para ver tus datos desde cualquier dispositivo.
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <Field label="Correo">
            <TextInput type="email" placeholder="tu@correo.com" value={email} onChange={(e) => setEmail(e.target.value)} />
          </Field>
          <Field label="Contraseña">
            <TextInput type="password" placeholder="••••••••" value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") submit(); }} />
          </Field>
          {error && <div style={{ fontSize: 12, color: BAD, background: `${BAD}12`, borderRadius: 8, padding: "8px 10px" }}>{error}</div>}
          <Button variant="primary" onClick={submit} disabled={loadingLogin} style={{ justifyContent: "center", marginTop: 4 }}>
            {loadingLogin ? "Entrando…" : "Iniciar sesión"}
          </Button>
          <button
            onClick={onOfflineMode}
            style={{ border: "none", background: "transparent", color: MUTED, fontSize: 12, cursor: "pointer", marginTop: 4 }}
          >
            Continuar sin conectarme (solo en este chat)
          </button>
        </div>
      </Card>
    </div>
  );
}

/* ============================== FORMULARIO DE VENTA ============================== */

const emptyForm = (cat) => ({
  id: null,
  fecha: todayISO(),
  numeroPedido: "",
  cliente: "",
  empresa: "",
  telefono: "",
  correo: "",
  modelo: cat.modelos[0] || "",
  cantidadPiezas: "",
  metrosCuadrados: "",
  color: cat.colores[0] || "",
  acabado: cat.acabados[0] || "",
  estadoDestino: ESTADOS_MX[0],
  ciudadDestino: "",
  tipoEnvio: "Local",
  fletera: cat.fleteras[0] || "",
  conIVA: false,
  subtotal: "",
  totalConIVAInput: "",
  anticipoPct: 60,
  anticipoCapturado: "",
  formaPago: cat.formasPago[0] || "",
  fechaRealCobro: "",
  pagos: [],
  observaciones: "",
  etapa: "Prospecto",
  proximoSeguimiento: "",
  notaSeguimiento: "",
});

// Resumen de folio + historial de pagos + mini-formulario para registrar un nuevo pago.
// Se usa tanto cuando se reconoce un folio existente al capturar "Nuevo pedido"
// como dentro de "Editar pedido".
function PagosHistorial({ venta, onAddPago, locked }) {
  const subtotal = getSubtotalPedido(venta);
  const pagos = useMemo(() => [...getPagosPedido(venta)].sort((a, b) => (a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0)), [venta]);
  const totalPagado = getTotalPagado(venta);
  const saldo = getSaldoPendiente(venta);
  const liquidado = getLiquidado(venta);
  // El pedido "involucra IVA" si así se marcó al crearlo, o si alguno de sus pagos se recibió con IVA.
  const pedidoTieneIVA = !!venta.conIVA || pagos.some((p) => p.conIVA) || Number(venta.totalCobrado) > 0;

  const [fecha, setFecha] = useState(todayISO());
  const [conIVAPago, setConIVAPago] = useState(!!venta.conIVA);
  const [monto, setMonto] = useState("");
  const [tipo, setTipo] = useState("Pago parcial");

  const montoRecibidoNum = Number(monto) || 0;
  const montoSubtotalPago = conIVAPago ? montoRecibidoNum / 1.16 : montoRecibidoNum;

  useEffect(() => {
    if (montoSubtotalPago > 0 && montoSubtotalPago >= saldo - 0.5) setTipo("Finiquito");
    else if (montoSubtotalPago > 0) setTipo((t) => (t === "Finiquito" ? "Pago parcial" : t));
  }, [montoSubtotalPago, saldo]);

  const excede = montoSubtotalPago > saldo + 0.01;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      {pedidoTieneIVA && (
        <div style={{ background: `${ACCENT}0d`, borderRadius: 10, padding: "12px 14px", display: "flex", gap: 22, flexWrap: "wrap" }}>
          <div style={{ fontSize: 10.5, color: ACCENT, fontWeight: 800, textTransform: "uppercase", width: "100%" }}>Con IVA (para cotejar con oficina)</div>
          <div><div style={{ fontSize: 11, color: MUTED, fontWeight: 600 }}>TOTAL CON IVA</div><div style={{ fontWeight: 700, color: INK }}>{fmtMoney(subtotal * 1.16)}</div></div>
          <div><div style={{ fontSize: 11, color: MUTED, fontWeight: 600 }}>PAGADO CON IVA</div><div style={{ fontWeight: 700, color: GOOD }}>{fmtMoney(totalPagado * 1.16)}</div></div>
          <div><div style={{ fontSize: 11, color: MUTED, fontWeight: 600 }}>SALDO CON IVA</div><div style={{ fontWeight: 700, color: liquidado ? GOOD : WARN }}>{fmtMoney(Math.max(0, saldo) * 1.16)}</div></div>
        </div>
      )}
      <div style={{ background: PAPER, borderRadius: 10, padding: "12px 14px", display: "flex", gap: 22, flexWrap: "wrap" }}>
        {pedidoTieneIVA && <div style={{ fontSize: 10.5, color: MUTED, fontWeight: 800, textTransform: "uppercase", width: "100%" }}>Subtotal (lo que cuenta para tu comisión)</div>}
        <div><div style={{ fontSize: 11, color: MUTED, fontWeight: 600 }}>{pedidoTieneIVA ? "SUBTOTAL" : "VALOR"} (IMPORTE ORIGINAL)</div><div style={{ fontWeight: 700, color: INK }}>{fmtMoney(subtotal)}</div></div>
        <div><div style={{ fontSize: 11, color: MUTED, fontWeight: 600 }}>TOTAL PAGADO</div><div style={{ fontWeight: 700, color: GOOD }}>{fmtMoney(totalPagado)}</div></div>
        <div>
          <div style={{ fontSize: 11, color: MUTED, fontWeight: 600 }}>SALDO PENDIENTE</div>
          {liquidado ? <Pill color={GOOD}>Liquidado</Pill> : <div style={{ fontWeight: 700, color: WARN }}>{fmtMoney(saldo)}</div>}
        </div>
      </div>

      {pagos.length > 0 && (
        <div style={{ border: `1px solid ${LINE}`, borderRadius: 10, overflow: "hidden" }}>
          <div style={{ padding: "6px 12px", fontSize: 10.5, color: MUTED, fontWeight: 700, textTransform: "uppercase", background: PAPER }}>Historial de movimientos</div>
          {pagos.map((p) => (
            <div key={p.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 12px", fontSize: 12.5, borderTop: `1px solid ${LINE}` }}>
              <span style={{ color: MUTED }}>{fmtDate(p.fecha)}</span>
              <Pill color={p.tipo === "Finiquito" ? GOOD : p.tipo === "Anticipo" ? ACCENT : WARN}>{p.tipo}</Pill>
              <span style={{ textAlign: "right" }}>
                <div style={{ fontWeight: 700 }}>{fmtMoney(p.conIVA ? (p.montoRecibido != null ? p.montoRecibido : p.monto * 1.16) : p.monto)}{p.conIVA && <span style={{ fontSize: 10.5, color: MUTED, fontWeight: 500 }}> con IVA</span>}</div>
                {p.conIVA && <div style={{ fontSize: 11, color: MUTED }}>Subtotal: {fmtMoney(p.monto)}</div>}
              </span>
            </div>
          ))}
        </div>
      )}

      {!locked && !liquidado && (
        <div style={{ border: `1px dashed ${LINE}`, borderRadius: 10, padding: 12, display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8 }}>
          <Field label="Fecha del pago"><TextInput type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
          <Field label="¿Este pago incluye IVA?">
            <Select value={conIVAPago ? "si" : "no"} onChange={(e) => setConIVAPago(e.target.value === "si")}>
              <option value="no">No, sin IVA</option>
              <option value="si">Sí, con IVA (transferencia)</option>
            </Select>
          </Field>
          <Field label="Tipo de movimiento">
            <Select value={tipo} onChange={(e) => setTipo(e.target.value)}>
              {TIPOS_MOVIMIENTO.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
          <Field label={conIVAPago ? "Monto recibido (con IVA)" : "Monto recibido"} span={2}>
            <TextInput type="number" min="0" placeholder="0.00" value={monto} onChange={(e) => setMonto(e.target.value)} />
          </Field>
          <div style={{ alignSelf: "end", fontSize: 12, color: MUTED, paddingBottom: 8 }}>
            {conIVAPago && montoRecibidoNum > 0 && (
              <>De estos <b style={{ color: INK }}>{fmtMoney(montoRecibidoNum)}</b>, <b style={{ color: INK }}>{fmtMoney(montoSubtotalPago)}</b> son subtotal.</>
            )}
          </div>
          {excede && (
            <div style={{ gridColumn: "span 3", fontSize: 12, color: BAD, background: `${BAD}12`, borderRadius: 8, padding: "6px 10px" }}>
              ⚠ El subtotal de este pago ({fmtMoney(montoSubtotalPago)}) supera el saldo pendiente ({fmtMoney(saldo)}).
            </div>
          )}
          <div style={{ gridColumn: "span 3" }}>
            <Button
              variant="primary" icon={Plus}
              disabled={!(montoRecibidoNum > 0 && fecha)}
              onClick={() => {
                if (excede && !window.confirm(`El subtotal de este pago (${fmtMoney(montoSubtotalPago)}) supera el saldo pendiente de ${fmtMoney(saldo)}. ¿Registrarlo de todas formas?`)) return;
                onAddPago({
                  id: uid(), fecha, tipo, conIVA: conIVAPago,
                  montoRecibido: round2(montoRecibidoNum), monto: round2(montoSubtotalPago),
                });
              }}
            >
              Registrar pago
            </Button>
          </div>
        </div>
      )}
    </div>
  );
}

function SaleModal({ initial, onClose, onSave, onAddPago, catalogos, clientesConocidos = [], ventasExistentes = [] }) {
  const [form, setForm] = useState(initial || emptyForm(catalogos));
  const set = (k) => (e) => {
    const val = e && e.target ? e.target.value : e;
    setForm((f) => ({ ...f, [k]: val }));
  };
  const setCliente = (e) => {
    const val = e.target.value;
    setForm((f) => {
      const match = clientesConocidos.find((c) => c.cliente.toLowerCase() === val.trim().toLowerCase());
      if (match) {
        return {
          ...f, cliente: val,
          empresa: f.empresa || match.empresa || "",
          telefono: f.telefono || match.telefono || "",
          correo: f.correo || match.correo || "",
        };
      }
      return { ...f, cliente: val };
    });
  };

  const esEdicion = !!initial;

  // Al escribir un folio en "Nuevo pedido", si ya existe se reconoce automáticamente.
  const folioCoincidente = !esEdicion && form.numeroPedido.trim()
    ? ventasExistentes.find((v) => (v.numeroPedido || "").trim().toLowerCase() === form.numeroPedido.trim().toLowerCase())
    : null;

  const subtotalNuevo = form.conIVA ? (Number(form.totalConIVAInput) || 0) / 1.16 : (Number(form.subtotal) || 0);
  const ivaNuevo = form.conIVA ? Math.max(0, (Number(form.totalConIVAInput) || 0) - subtotalNuevo) : 0;
  const anticipoSugerido = round2(subtotalNuevo * ((Number(form.anticipoPct) || 0) / 100));
  const anticipoRecibido = Number(form.anticipoCapturado) || 0;
  const anticipoNum = round2(form.conIVA ? anticipoRecibido / 1.16 : anticipoRecibido);
  const saldoNuevo = round2(subtotalNuevo - anticipoNum);

  const canSaveNuevo = form.cliente.trim() && form.numeroPedido.trim() && form.fecha;
  const canSaveEdicion = form.cliente.trim() && form.numeroPedido.trim() && form.fecha;

  const handleGuardarNuevo = () => {
    const cleanNum = (v) => (v === "" || v == null ? 0 : Number(v));
    const pagosIniciales = [];
    if (anticipoNum > 0 && form.fechaRealCobro) {
      pagosIniciales.push({
        id: uid(), fecha: form.fechaRealCobro, conIVA: !!form.conIVA,
        montoRecibido: round2(anticipoRecibido), monto: anticipoNum,
        tipo: anticipoNum >= subtotalNuevo - 0.5 ? "Finiquito" : "Anticipo",
      });
    }
    onSave({
      ...form,
      id: uid(),
      cantidadPiezas: cleanNum(form.cantidadPiezas),
      metrosCuadrados: cleanNum(form.metrosCuadrados),
      subtotal: subtotalNuevo > 0 ? round2(subtotalNuevo) : "",
      conIVA: !!form.conIVA,
      anticipoPct: Number(form.anticipoPct) || 60,
      pagos: pagosIniciales,
      etapa: pagosIniciales.length && anticipoNum >= subtotalNuevo - 0.5 ? "Cobrado" : form.etapa,
    });
  };

  const handleGuardarEdicion = () => {
    const cleanNum = (v) => (v === "" || v == null ? 0 : Number(v));
    onSave({
      ...form,
      id: initial.id,
      cantidadPiezas: cleanNum(form.cantidadPiezas),
      metrosCuadrados: cleanNum(form.metrosCuadrados),
    });
  };

  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(14,26,42,0.45)", zIndex: 60,
      display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "28px 16px", overflowY: "auto",
    }}>
      <Card style={{ width: "100%", maxWidth: 780, padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "16px 20px", borderBottom: `1px solid ${LINE}`, display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, background: "#fff", zIndex: 2 }}>
          <div style={{ fontWeight: 700, fontSize: 15.5, color: INK }}>
            {esEdicion ? "Editar pedido" : folioCoincidente ? "Folio existente" : "Nuevo pedido"}
          </div>
          <button onClick={onClose} style={{ border: "none", background: "transparent", cursor: "pointer", color: MUTED }}>
            <X size={20} />
          </button>
        </div>

        {folioCoincidente ? (
          /* ---------- Folio ya existente: datos bloqueados + solo captura de pago ---------- */
          <div style={{ padding: 20, display: "flex", flexDirection: "column", gap: 14, maxHeight: "70vh", overflowY: "auto" }}>
            <div style={{ fontSize: 12.5, color: ACCENT, background: `${ACCENT}12`, borderRadius: 8, padding: "9px 12px" }}>
              El pedido <b>{folioCoincidente.numeroPedido}</b> ya existe. No se crea un pedido nuevo — solo registra el pago.
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 10, fontSize: 13 }}>
              <div><div style={{ fontSize: 11, color: MUTED, fontWeight: 700, textTransform: "uppercase" }}>Cliente</div><div style={{ fontWeight: 600 }}>{folioCoincidente.cliente}</div></div>
              <div><div style={{ fontSize: 11, color: MUTED, fontWeight: 700, textTransform: "uppercase" }}>Empresa</div><div>{folioCoincidente.empresa || "—"}</div></div>
              <div><div style={{ fontSize: 11, color: MUTED, fontWeight: 700, textTransform: "uppercase" }}>Modelo / cantidades</div><div>{folioCoincidente.modelo} · {fmtNum(folioCoincidente.cantidadPiezas)} pzas · {fmtNum(folioCoincidente.metrosCuadrados)} m²</div></div>
              <div><div style={{ fontSize: 11, color: MUTED, fontWeight: 700, textTransform: "uppercase" }}>Destino</div><div>{folioCoincidente.ciudadDestino}, {folioCoincidente.estadoDestino}</div></div>
            </div>
            <PagosHistorial
              venta={folioCoincidente}
              onAddPago={(pago) => { onAddPago(folioCoincidente.id, pago); onClose(); }}
            />
          </div>
        ) : (
          <>
            <div style={{ padding: 20, display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 12, maxHeight: "65vh", overflowY: "auto" }}>
              <Field label="Fecha"><TextInput type="date" value={form.fecha} onChange={set("fecha")} /></Field>
              <Field label="Número de pedido (folio)">
                <TextInput placeholder="PED-1001" value={form.numeroPedido} onChange={set("numeroPedido")} disabled={esEdicion} />
              </Field>

              <Field label="Cliente">
                <TextInput list="clientes-datalist" placeholder="Escribe o elige un cliente existente" value={form.cliente} onChange={setCliente} />
                <datalist id="clientes-datalist">
                  {clientesConocidos.map((c) => <option key={c.cliente} value={c.cliente} />)}
                </datalist>
              </Field>
              <Field label="Empresa"><TextInput placeholder="Razón social" value={form.empresa} onChange={set("empresa")} /></Field>

              <Field label="Teléfono"><TextInput placeholder="999 000 0000" value={form.telefono} onChange={set("telefono")} /></Field>
              <Field label="Correo"><TextInput type="email" placeholder="cliente@correo.com" value={form.correo} onChange={set("correo")} /></Field>

              <Field label="Modelo">
                <Select value={form.modelo} onChange={set("modelo")}>
                  {catalogos.modelos.map((m) => <option key={m}>{m}</option>)}
                </Select>
              </Field>
              <Field label="Etapa">
                <Select value={form.etapa} onChange={set("etapa")}>
                  {ETAPAS.map((e) => <option key={e}>{e}</option>)}
                </Select>
              </Field>

              <Field label="Cantidad de piezas"><TextInput type="number" min="0" value={form.cantidadPiezas} onChange={set("cantidadPiezas")} /></Field>
              <Field label="Metros cuadrados (m²)"><TextInput type="number" min="0" value={form.metrosCuadrados} onChange={set("metrosCuadrados")} /></Field>

              <Field label="Color / Paleta">
                <Select value={form.color} onChange={set("color")}>{catalogos.colores.map((c) => <option key={c}>{c}</option>)}</Select>
              </Field>
              <Field label="Acabado">
                <Select value={form.acabado} onChange={set("acabado")}>{catalogos.acabados.map((a) => <option key={a}>{a}</option>)}</Select>
              </Field>

              <Field label="Estado destino">
                <Select value={form.estadoDestino} onChange={set("estadoDestino")}>
                  {ESTADOS_MX.map((e) => <option key={e}>{e}</option>)}
                </Select>
              </Field>
              <Field label="Ciudad destino"><TextInput placeholder="Ciudad" value={form.ciudadDestino} onChange={set("ciudadDestino")} /></Field>

              <Field label="Tipo de envío">
                <Select value={form.tipoEnvio} onChange={set("tipoEnvio")}>
                  <option>Local</option><option>Nacional</option>
                </Select>
              </Field>
              <Field label="Fletera">
                <Select value={form.fletera} onChange={set("fletera")}>{catalogos.fleteras.map((f) => <option key={f}>{f}</option>)}</Select>
              </Field>

              <Field label="Forma de pago">
                <Select value={form.formaPago} onChange={set("formaPago")}>{catalogos.formasPago.map((f) => <option key={f}>{f}</option>)}</Select>
              </Field>
              <Field label="Próximo seguimiento (recordatorio)"><TextInput type="datetime-local" value={form.proximoSeguimiento} onChange={set("proximoSeguimiento")} /></Field>

              <Field label="¿Qué hay que hacer?" span={2}><TextInput placeholder="Ej. Llamarle para confirmar precio" value={form.notaSeguimiento} onChange={set("notaSeguimiento")} /></Field>

              <Field label="Observaciones" span={2}>
                <TextArea placeholder="Notas del pedido…" value={form.observaciones} onChange={set("observaciones")} />
              </Field>

              {/* ---------- Bloque de dinero ---------- */}
              {esEdicion ? (
                <div style={{ gridColumn: "span 2", display: "flex", flexDirection: "column", gap: 10 }}>
                  <div style={{ fontSize: 11, color: MUTED, fontWeight: 700, textTransform: "uppercase" }}>Cobro del pedido (folio bloqueado)</div>
                  <PagosHistorial venta={initial} onAddPago={(pago) => { onAddPago(initial.id, pago); onClose(); }} />
                </div>
              ) : (
                <>
                  <Field label="¿El monto que capturas incluye IVA?">
                    <Select value={form.conIVA ? "si" : "no"} onChange={(e) => set("conIVA")(e.target.value === "si")}>
                      <option value="no">Sin IVA (yo capturo el subtotal)</option>
                      <option value="si">Con IVA (yo capturo el total)</option>
                    </Select>
                  </Field>
                  {form.conIVA ? (
                    <Field label="Total con IVA"><TextInput type="number" min="0" placeholder="0.00" value={form.totalConIVAInput} onChange={set("totalConIVAInput")} /></Field>
                  ) : (
                    <Field label="Subtotal (sin IVA)"><TextInput type="number" min="0" placeholder="0.00" value={form.subtotal} onChange={set("subtotal")} /></Field>
                  )}

                  <Field label="% anticipo sugerido (editable)"><TextInput type="number" min="0" max="100" value={form.anticipoPct} onChange={set("anticipoPct")} /></Field>
                  <Field label="Fecha real de cobro (del anticipo)"><TextInput type="date" value={form.fechaRealCobro} onChange={set("fechaRealCobro")} /></Field>

                  <Field label={form.conIVA ? "Anticipo capturado (con IVA)" : "Anticipo capturado"}>
                    <TextInput type="number" min="0" placeholder={anticipoSugerido ? String(round2(form.conIVA ? anticipoSugerido * 1.16 : anticipoSugerido)) : "0.00"} value={form.anticipoCapturado} onChange={set("anticipoCapturado")} />
                  </Field>
                  <div style={{ alignSelf: "end", fontSize: 12, color: MUTED, paddingBottom: 8 }}>
                    Sugerido ({form.anticipoPct || 0}%): <b style={{ color: INK }}>{fmtMoney(form.conIVA ? anticipoSugerido * 1.16 : anticipoSugerido)}</b>
                    {form.conIVA && anticipoRecibido > 0 && <> · de eso, <b style={{ color: INK }}>{fmtMoney(anticipoNum)}</b> son subtotal</>}
                  </div>

                  {form.conIVA && (
                    <div style={{ gridColumn: "span 2", background: `${ACCENT}0d`, borderRadius: 10, padding: "10px 14px", display: "flex", gap: 22, flexWrap: "wrap" }}>
                      <div style={{ fontSize: 10.5, color: ACCENT, fontWeight: 800, textTransform: "uppercase", width: "100%" }}>Con IVA (para cotejar con oficina)</div>
                      <div><div style={{ fontSize: 11, color: MUTED, fontWeight: 600 }}>TOTAL CON IVA</div><div style={{ fontWeight: 700, color: INK }}>{fmtMoney(subtotalNuevo * 1.16)}</div></div>
                      <div><div style={{ fontSize: 11, color: MUTED, fontWeight: 600 }}>ANTICIPO CON IVA</div><div style={{ fontWeight: 700, color: GOOD }}>{fmtMoney(anticipoRecibido)}</div></div>
                      <div><div style={{ fontSize: 11, color: MUTED, fontWeight: 600 }}>SALDO CON IVA</div><div style={{ fontWeight: 700, color: saldoNuevo > 0 ? WARN : GOOD }}>{fmtMoney(Math.max(0, saldoNuevo) * 1.16)}</div></div>
                    </div>
                  )}
                  <div style={{ gridColumn: "span 2", background: PAPER, borderRadius: 10, padding: "10px 14px", display: "flex", gap: 22, flexWrap: "wrap" }}>
                    {form.conIVA && <div style={{ fontSize: 10.5, color: MUTED, fontWeight: 800, textTransform: "uppercase", width: "100%" }}>Subtotal (lo que cuenta para tu comisión)</div>}
                    <div>
                      <div style={{ fontSize: 11, color: MUTED, fontWeight: 600 }}>{form.conIVA ? "SUBTOTAL" : "VALOR"}</div>
                      <div style={{ fontWeight: 700, color: INK }}>{fmtMoney(subtotalNuevo)}</div>
                    </div>
                    <div>
                      <div style={{ fontSize: 11, color: MUTED, fontWeight: 600 }}>{form.conIVA ? "ANTICIPO SUBTOTAL" : "ANTICIPO"}</div>
                      <div style={{ fontWeight: 700, color: GOOD }}>{fmtMoney(anticipoNum)}</div>
                    </div>
                    <div>
                      <div style={{ fontSize: 11, color: MUTED, fontWeight: 600 }}>SALDO PENDIENTE</div>
                      <div style={{ fontWeight: 700, color: saldoNuevo > 0 ? WARN : GOOD }}>{fmtMoney(Math.max(0, saldoNuevo))}</div>
                    </div>
                  </div>
                  <div style={{ gridColumn: "span 2", fontSize: 11.5, color: MUTED }}>
                    La comisión siempre se calcula sobre el subtotal sin IVA, nunca sobre el total con IVA.
                  </div>
                </>
              )}
            </div>

            <div style={{ padding: "14px 20px", borderTop: `1px solid ${LINE}`, display: "flex", justifyContent: "flex-end", gap: 8 }}>
              <Button variant="secondary" onClick={onClose}>Cancelar</Button>
              <Button
                variant="primary"
                icon={CheckCircle2}
                disabled={esEdicion ? !canSaveEdicion : !canSaveNuevo}
                onClick={esEdicion ? handleGuardarEdicion : handleGuardarNuevo}
              >
                {esEdicion ? "Guardar cambios" : "Guardar pedido"}
              </Button>
            </div>
          </>
        )}
      </Card>
    </div>
  );
}

/* ============================== COTIZACIÓN IMPRIMIBLE ============================== */

function QuoteModal({ venta, empresaInfo, onClose }) {
  const subtotal = getSubtotalPedido(venta);
  const iva = venta.conIVA ? subtotal * 0.16 : 0;
  const total = venta.conIVA ? subtotal + iva : subtotal;
  const nombreEmpresa = (empresaInfo && empresaInfo.nombre) || "Tu empresa";

  return (
    <div style={{
      position: "fixed", inset: 0, background: "#fff", zIndex: 80,
      display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "28px 16px", overflowY: "auto",
    }}>
      <Card style={{ width: "100%", maxWidth: 640, padding: 0, overflow: "hidden", border: "none", boxShadow: "0 6px 30px rgba(14,42,71,0.12)" }}>
        <div className="no-print" style={{ padding: "14px 20px", borderBottom: `1px solid ${LINE}`, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ fontWeight: 700, fontSize: 15.5, color: INK }}>Vista previa de cotización</div>
          <button onClick={onClose} style={{ border: "none", background: "transparent", cursor: "pointer", color: MUTED }}><X size={20} /></button>
        </div>

        <div style={{ padding: 28 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 22, borderBottom: `2px solid ${NAVY}`, paddingBottom: 16 }}>
            <div>
              <div style={{ fontWeight: 800, fontSize: 19, color: NAVY }}>{nombreEmpresa}</div>
              {empresaInfo?.telefono && <div style={{ fontSize: 12.5, color: MUTED }}>{empresaInfo.telefono}</div>}
              {empresaInfo?.correo && <div style={{ fontSize: 12.5, color: MUTED }}>{empresaInfo.correo}</div>}
            </div>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontWeight: 800, fontSize: 15, color: INK }}>COTIZACIÓN</div>
              <div style={{ fontSize: 12.5, color: MUTED }}>Pedido {venta.numeroPedido}</div>
              <div style={{ fontSize: 12.5, color: MUTED }}>{fmtDate(venta.fecha)}</div>
            </div>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 22, fontSize: 13 }}>
            <div>
              <div style={{ fontSize: 11, color: MUTED, fontWeight: 700, textTransform: "uppercase" }}>Cliente</div>
              <div style={{ fontWeight: 600 }}>{venta.cliente}</div>
              <div style={{ color: MUTED }}>{venta.empresa}</div>
              <div style={{ color: MUTED }}>{venta.telefono} {venta.correo && `· ${venta.correo}`}</div>
            </div>
            <div>
              <div style={{ fontSize: 11, color: MUTED, fontWeight: 700, textTransform: "uppercase" }}>Destino</div>
              <div>{venta.ciudadDestino}, {venta.estadoDestino}</div>
              <div style={{ color: MUTED }}>Envío {venta.tipoEnvio}{venta.fletera ? ` · ${venta.fletera}` : ""}</div>
            </div>
          </div>

          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, marginBottom: 18 }}>
            <thead>
              <tr style={{ background: PAPER }}>
                {["Modelo", "Color", "Acabado", "Piezas", "m²"].map((h) => (
                  <th key={h} style={{ textAlign: "left", padding: "8px 10px", fontSize: 11, color: MUTED, fontWeight: 700, textTransform: "uppercase" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              <tr style={{ borderTop: `1px solid ${LINE}` }}>
                <td style={{ padding: "9px 10px", fontWeight: 600 }}>{venta.modelo}</td>
                <td style={{ padding: "9px 10px" }}>{venta.color}</td>
                <td style={{ padding: "9px 10px" }}>{venta.acabado}</td>
                <td style={{ padding: "9px 10px" }}>{fmtNum(venta.cantidadPiezas)}</td>
                <td style={{ padding: "9px 10px" }}>{fmtNum(venta.metrosCuadrados)}</td>
              </tr>
            </tbody>
          </table>

          <div style={{ display: "flex", justifyContent: "flex-end" }}>
            <div style={{ width: 220, fontSize: 13.5 }}>
              <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 0" }}><span style={{ color: MUTED }}>Subtotal</span><span>{fmtMoney(subtotal)}</span></div>
              <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 0" }}><span style={{ color: MUTED }}>IVA (16%)</span><span>{fmtMoney(iva)}</span></div>
              <div style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderTop: `1px solid ${LINE}`, marginTop: 4, fontWeight: 800, fontSize: 15 }}><span>Total</span><span>{fmtMoney(total)}</span></div>
            </div>
          </div>

          {venta.observaciones && (
            <div style={{ marginTop: 18, fontSize: 12.5, color: MUTED }}>
              <div style={{ fontWeight: 700, color: INK, fontSize: 11, textTransform: "uppercase", marginBottom: 4 }}>Notas</div>
              {venta.observaciones}
            </div>
          )}
        </div>

        <div className="no-print" style={{ padding: "14px 20px", borderTop: `1px solid ${LINE}`, display: "flex", justifyContent: "flex-end", gap: 8 }}>
          <Button variant="secondary" onClick={onClose}>Cerrar</Button>
          <Button icon={Printer} onClick={() => printClean(`Cotización ${venta.numeroPedido || ""}`)}>Imprimir / Guardar PDF</Button>
        </div>
      </Card>
    </div>
  );
}

/* ==================================================================================
   MÓDULO CRM — completamente independiente del módulo de ventas/comisiones.
   No comparte estado, storage, ni componentes con SaleModal / KanbanBoard / ventas.
   ================================================================================== */

const CRM_ETAPAS = [
  "Oportunidad", "Cotización enviada", "Seguimiento", "Negociación",
  "Pedido confirmado", "En proceso", "Inactivo / Perdido",
];
const CRM_ETAPA_COLOR = {
  "Oportunidad": "#8C99AB",
  "Cotización enviada": "#3E7CB8",
  "Seguimiento": "#C88A2C",
  "Negociación": "#7A4FC2",
  "Pedido confirmado": "#2C8F5B",
  "En proceso": "#2C8FAE",
  "Inactivo / Perdido": "#9AA3AF",
};
const CRM_PRIORIDADES = ["Alta", "Media", "Baja"];
const CRM_PRIORIDAD_COLOR = { Alta: BAD, Media: WARN, Baja: MUTED };

const DEFAULT_ACCIONES_CRM = [
  "Solicitar información", "Solicitar medidas", "Solicitar dirección", "Cotizar flete",
  "Elaborar cotización", "Enviar cotización", "Dar seguimiento", "Llamar al cliente",
  "Negociar", "Esperar anticipo", "Confirmar producción", "Confirmar entrega",
];

const CRM_STORAGE_KEY = "crm-oportunidades-v1";
const CRM_ACCIONES_KEY = "crm-acciones-v1";
const CRM_ENFRIANDOSE_DIAS = 10;

async function loadCRMOportunidades() {
  if (currentSession) {
    try {
      const rows = await supaSelect("oportunidades");
      if (rows.length > 0) {
        const ops = rows.map(rowToOportunidad);
        await storage.set(CRM_STORAGE_KEY, JSON.stringify(ops), false).catch(() => {});
        return ops;
      }
      const localRes = await storage.get(CRM_STORAGE_KEY, false).catch(() => null);
      const local = localRes && localRes.value ? JSON.parse(localRes.value) : null;
      if (Array.isArray(local) && local.length) {
        await supaReplaceTable("oportunidades", local.map(oportunidadToRow)).catch((e) => console.error("No se pudo subir a Supabase", e));
        return local;
      }
      return [];
    } catch (e) {
      console.error("Fallo Supabase, usando copia local", e);
    }
  }
  try {
    const res = await storage.get(CRM_STORAGE_KEY, false);
    if (res && res.value) return JSON.parse(res.value);
    return null;
  } catch (e) { return null; }
}
async function saveCRMOportunidades(data) {
  try { await storage.set(CRM_STORAGE_KEY, JSON.stringify(data), false); }
  catch (e) { console.error("No se pudo guardar localmente", e); }
  if (currentSession) {
    try { await supaReplaceTable("oportunidades", data.map(oportunidadToRow)); }
    catch (e) { console.error("No se pudo sincronizar el CRM con Supabase", e); return false; }
  }
  return true;
}
async function loadCRMAcciones() {
  if (currentSession) {
    try {
      const res = await supaFetch(`${SUPABASE_URL}/rest/v1/app_config?key=eq.crm_acciones&select=value`);
      if (res.ok) {
        const rows = await res.json();
        if (rows.length) {
          await storage.set(CRM_ACCIONES_KEY, JSON.stringify(rows[0].value), false).catch(() => {});
          return rows[0].value;
        }
        const localRes = await storage.get(CRM_ACCIONES_KEY, false).catch(() => null);
        const local = localRes && localRes.value ? JSON.parse(localRes.value) : null;
        if (Array.isArray(local) && local.length) {
          await supaUpsert("app_config", { key: "crm_acciones", value: local }).catch((e) => console.error("No se pudo subir a Supabase", e));
          return local;
        }
      }
    } catch (e) {
      console.error("Fallo Supabase, usando copia local", e);
    }
  }
  try {
    const res = await storage.get(CRM_ACCIONES_KEY, false);
    if (res && res.value) return JSON.parse(res.value);
    return null;
  } catch (e) { return null; }
}
async function saveCRMAcciones(data) {
  try { await storage.set(CRM_ACCIONES_KEY, JSON.stringify(data), false); }
  catch (e) { console.error("No se pudo guardar localmente", e); }
  if (currentSession) {
    try { await supaUpsert("app_config", { key: "crm_acciones", value: data }); }
    catch (e) { console.error("No se pudo sincronizar acciones del CRM con Supabase", e); return false; }
  }
  return true;
}

function crmSeedData() {
  const today = new Date();
  const iso = (offsetDays) => {
    const d = new Date(today);
    d.setDate(d.getDate() + offsetDays);
    return d.toISOString().slice(0, 10);
  };
  const isoHora = (offsetDays, hora) => `${iso(offsetDays)}T${String(hora).padStart(2, "0")}:00`;
  return [
    {
      id: uid(), cliente: "Karla Solís", empresa: "", ciudad: "Mérida", telefono: "9991110022",
      valorEstimado: 65000, fechaUltimoContacto: iso(-2), proximaAccion: "Elaborar cotización",
      fechaProximaAccion: isoHora(0, 16), prioridad: "Alta", notas: "Quiere celosía para fachada, ya mandó medidas.",
      etapa: "Oportunidad", creadoEn: iso(-2), seguimientos: [],
    },
    {
      id: uid(), cliente: "Iván Chan", empresa: "Chan Arquitectos", ciudad: "Cancún", telefono: "9982223344",
      valorEstimado: 180000, fechaUltimoContacto: iso(-5), proximaAccion: "Dar seguimiento",
      fechaProximaAccion: isoHora(-1, 11), prioridad: "Alta", notas: "Proyecto de fraccionamiento, esperando aprobación de presupuesto.",
      etapa: "Cotización enviada", creadoEn: iso(-9),
      seguimientos: [
        { id: uid(), fecha: isoHora(-9, 10), nota: "Primer contacto, se envió catálogo." },
        { id: uid(), fecha: isoHora(-5, 12), nota: "Confirmó que le interesa, va a revisar presupuesto." },
      ],
    },
    {
      id: uid(), cliente: "Paola Briceño", empresa: "", ciudad: "Playa del Carmen", telefono: "9843335566",
      valorEstimado: 42000, fechaUltimoContacto: iso(-12), proximaAccion: "Llamar al cliente",
      fechaProximaAccion: isoHora(-3, 9), prioridad: "Media", notas: "Se enfrió después de la cotización, hay que retomar.",
      etapa: "Seguimiento", creadoEn: iso(-18),
      seguimientos: [
        { id: uid(), fecha: isoHora(-18, 9), nota: "Solicitó cotización." },
        { id: uid(), fecha: isoHora(-15, 10), nota: "Se envió cotización por correo." },
        { id: uid(), fecha: isoHora(-12, 16), nota: "No contestó llamada." },
      ],
    },
    {
      id: uid(), cliente: "Grupo Osorno", empresa: "Constructora Osorno", ciudad: "Campeche", telefono: "9814445577",
      valorEstimado: 310000, fechaUltimoContacto: iso(-1), proximaAccion: "Negociar",
      fechaProximaAccion: isoHora(1, 12), prioridad: "Alta", notas: "Piden descuento por volumen, esperando autorización.",
      etapa: "Negociación", creadoEn: iso(-14), seguimientos: [{ id: uid(), fecha: isoHora(-14, 9), nota: "Primer contacto, mandó planos del proyecto." }],
    },
    {
      id: uid(), cliente: "Rodrigo Balam", empresa: "", ciudad: "Valladolid", telefono: "9857778899",
      valorEstimado: 28000, fechaUltimoContacto: iso(-20), proximaAccion: "Dar seguimiento",
      fechaProximaAccion: isoHora(-8, 10), prioridad: "Baja", notas: "No contestó las últimas 2 llamadas.",
      etapa: "Inactivo / Perdido", creadoEn: iso(-30), seguimientos: [{ id: uid(), fecha: isoHora(-30, 11), nota: "Primer contacto." }, { id: uid(), fecha: isoHora(-20, 15), nota: "Segundo intento, no contestó." }],
    },
  ];
}

const emptyOpportunity = (acciones) => ({
  id: null,
  cliente: "",
  empresa: "",
  ciudad: "",
  telefono: "",
  valorEstimado: "",
  fechaUltimoContacto: todayISO(),
  proximaAccion: acciones[0] || "",
  fechaProximaAccion: nowLocalISO(),
  prioridad: "Media",
  notas: "",
  etapa: "Oportunidad",
  creadoEn: todayISO(),
  seguimientos: [],
  modelo: "",
  motivoPerdida: "",
});

function OpportunityModal({ initial, onClose, onSave, onDelete, acciones, onAddAccion }) {
  const esPersonalizadaInicial = !!(initial && initial.proximaAccion && !acciones.includes(initial.proximaAccion));
  const [form, setForm] = useState(initial || emptyOpportunity(acciones));
  const [usarPersonalizada, setUsarPersonalizada] = useState(esPersonalizadaInicial);
  const [accionPersonalizada, setAccionPersonalizada] = useState(esPersonalizadaInicial ? initial.proximaAccion : "");

  const [mostrarSeguimiento, setMostrarSeguimiento] = useState(false);
  const [notaSeguimientoNuevo, setNotaSeguimientoNuevo] = useState("");
  const [accionSeguimientoNuevo, setAccionSeguimientoNuevo] = useState(form.proximaAccion || acciones[0] || "");
  const [fechaCustomSeguimiento, setFechaCustomSeguimiento] = useState("");
  const [mostrarCustom, setMostrarCustom] = useState(false);

  const set = (k) => (e) => {
    const val = e && e.target ? e.target.value : e;
    setForm((f) => ({ ...f, [k]: val }));
  };

  const canSave = form.cliente.trim() && (!usarPersonalizada || accionPersonalizada.trim());

  const registrarSeguimiento = (opcion) => {
    const ahora = nowLocalISO();
    const nuevoSeg = { id: uid(), fecha: ahora, nota: notaSeguimientoNuevo.trim() };
    let nuevaFechaAccion;
    if (opcion === "manana") nuevaFechaAccion = addDaysLocalISO(1);
    else if (opcion === "3dias") nuevaFechaAccion = addDaysLocalISO(3);
    else if (opcion === "1semana") nuevaFechaAccion = addDaysLocalISO(7);
    else nuevaFechaAccion = fechaCustomSeguimiento || form.fechaProximaAccion;

    const actualizado = {
      ...form,
      id: form.id || uid(),
      seguimientos: [...(form.seguimientos || []), nuevoSeg],
      fechaUltimoContacto: todayISO(),
      proximaAccion: accionSeguimientoNuevo || form.proximaAccion,
      fechaProximaAccion: nuevaFechaAccion,
      valorEstimado: form.valorEstimado === "" ? 0 : Number(form.valorEstimado),
    };
    onSave(actualizado);
  };

  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(14,26,42,0.45)", zIndex: 60,
      display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "28px 16px", overflowY: "auto",
    }}>
      <Card style={{ width: "100%", maxWidth: 640, padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "16px 20px", borderBottom: `1px solid ${LINE}`, display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, background: "#fff", zIndex: 2 }}>
          <div style={{ fontWeight: 700, fontSize: 15.5, color: INK }}>
            {initial ? "Editar oportunidad" : "Nueva oportunidad"}
          </div>
          <button onClick={onClose} style={{ border: "none", background: "transparent", cursor: "pointer", color: MUTED }}>
            <X size={20} />
          </button>
        </div>

        <div style={{ padding: 20, display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 12, maxHeight: "65vh", overflowY: "auto" }}>
          <Field label="Nombre del cliente"><TextInput placeholder="Nombre" value={form.cliente} onChange={set("cliente")} /></Field>
          <Field label="Empresa (opcional)"><TextInput placeholder="Empresa" value={form.empresa} onChange={set("empresa")} /></Field>

          <Field label="Ciudad"><TextInput placeholder="Ciudad" value={form.ciudad} onChange={set("ciudad")} /></Field>
          <Field label="Teléfono"><TextInput placeholder="999 000 0000" value={form.telefono} onChange={set("telefono")} /></Field>

          <Field label="Modelo / detalle (opcional)"><TextInput placeholder="Ej. C2005, 15 m²" value={form.modelo} onChange={set("modelo")} /></Field>
          <Field label="Valor estimado"><TextInput type="number" min="0" placeholder="0.00" value={form.valorEstimado} onChange={set("valorEstimado")} /></Field>

          <Field label="Prioridad">
            <Select value={form.prioridad} onChange={set("prioridad")}>
              {CRM_PRIORIDADES.map((p) => <option key={p}>{p}</option>)}
            </Select>
          </Field>
          <Field label="Etapa">
            <Select value={form.etapa} onChange={set("etapa")}>
              {CRM_ETAPAS.map((e) => <option key={e}>{e}</option>)}
            </Select>
          </Field>

          {form.etapa === "Inactivo / Perdido" && (
            <Field label="Motivo de pérdida" span={2}>
              <TextInput placeholder="Ej. Precio, se fue con otro proveedor, dejó de contestar…" value={form.motivoPerdida} onChange={set("motivoPerdida")} />
            </Field>
          )}

          <Field label="Fecha del último contacto"><TextInput type="date" value={form.fechaUltimoContacto} onChange={set("fechaUltimoContacto")} /></Field>

          <Field label="Próxima acción">
            {!usarPersonalizada ? (
              <Select
                value={form.proximaAccion}
                onChange={(e) => {
                  if (e.target.value === "__custom__") { setUsarPersonalizada(true); setAccionPersonalizada(""); }
                  else set("proximaAccion")(e);
                }}
              >
                {acciones.map((a) => <option key={a}>{a}</option>)}
                <option value="__custom__">+ Otra (personalizada)…</option>
              </Select>
            ) : (
              <div style={{ display: "flex", gap: 6 }}>
                <TextInput autoFocus placeholder="Escribe la acción…" value={accionPersonalizada} onChange={(e) => setAccionPersonalizada(e.target.value)} />
                <button
                  type="button"
                  onClick={() => { setUsarPersonalizada(false); setAccionPersonalizada(""); }}
                  title="Cancelar, usar lista"
                  style={{ border: `1px solid ${LINE}`, background: "#fff", borderRadius: 9, padding: "0 10px", cursor: "pointer", color: MUTED }}
                >
                  <X size={14} />
                </button>
              </div>
            )}
          </Field>
          <Field label="Fecha y hora de la próxima acción"><TextInput type="datetime-local" value={form.fechaProximaAccion} onChange={set("fechaProximaAccion")} /></Field>

          <Field label="Notas" span={2}>
            <TextArea placeholder="Contexto de la oportunidad, lo que ya se habló, condiciones…" value={form.notas} onChange={set("notas")} />
          </Field>

          {initial && (
            <div style={{ gridColumn: "span 2", borderTop: `1px solid ${LINE}`, paddingTop: 14, marginTop: 4 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
                <div style={{ fontSize: 12.5, fontWeight: 700, color: INK }}>
                  Historial de seguimientos ({(form.seguimientos || []).length})
                </div>
                {!mostrarSeguimiento && (
                  <Button size="sm" icon={Plus} onClick={() => setMostrarSeguimiento(true)}>Registrar seguimiento</Button>
                )}
              </div>

              {(form.seguimientos || []).length > 0 && (
                <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: mostrarSeguimiento ? 12 : 0, maxHeight: 140, overflowY: "auto" }}>
                  {[...(form.seguimientos || [])].reverse().map((s) => (
                    <div key={s.id} style={{ fontSize: 12, background: PAPER, borderRadius: 8, padding: "6px 10px", display: "flex", gap: 8 }}>
                      <span style={{ color: MUTED, fontWeight: 600, flexShrink: 0 }}>{fmtDateTime(s.fecha)}</span>
                      <span style={{ color: INK }}>{s.nota || "Seguimiento registrado (sin nota)"}</span>
                    </div>
                  ))}
                </div>
              )}

              {mostrarSeguimiento && (
                <div style={{ background: PAPER, borderRadius: 10, padding: 12, display: "flex", flexDirection: "column", gap: 8 }}>
                  <TextInput
                    placeholder="¿Qué pasó? (opcional) ej. Le marqué, no contestó"
                    value={notaSeguimientoNuevo}
                    onChange={(e) => setNotaSeguimientoNuevo(e.target.value)}
                  />
                  <Field label="¿Próxima acción?">
                    <Select value={accionSeguimientoNuevo} onChange={(e) => setAccionSeguimientoNuevo(e.target.value)}>
                      {acciones.map((a) => <option key={a}>{a}</option>)}
                    </Select>
                  </Field>
                  <div style={{ fontSize: 11.5, color: MUTED }}>¿Cuándo le vuelvo a dar seguimiento?</div>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    <Button size="sm" variant="subtle" onClick={() => registrarSeguimiento("manana")}>Mañana</Button>
                    <Button size="sm" variant="subtle" onClick={() => registrarSeguimiento("3dias")}>En 3 días</Button>
                    <Button size="sm" variant="subtle" onClick={() => registrarSeguimiento("1semana")}>En 1 semana</Button>
                    <Button size="sm" variant="subtle" onClick={() => setMostrarCustom((v) => !v)}>Otra fecha</Button>
                    <Button size="sm" variant="ghost" onClick={() => { setMostrarSeguimiento(false); setNotaSeguimientoNuevo(""); setMostrarCustom(false); }}>Cancelar</Button>
                  </div>
                  {mostrarCustom && (
                    <div style={{ display: "flex", gap: 6 }}>
                      <TextInput type="datetime-local" value={fechaCustomSeguimiento} onChange={(e) => setFechaCustomSeguimiento(e.target.value)} />
                      <Button size="sm" disabled={!fechaCustomSeguimiento} onClick={() => registrarSeguimiento("otra")}>Confirmar</Button>
                    </div>
                  )}
                </div>
              )}
            </div>
          )}
        </div>

        <div style={{ padding: "14px 20px", borderTop: `1px solid ${LINE}`, display: "flex", justifyContent: "space-between", gap: 8 }}>
          {initial ? (
            <Button variant="danger" icon={Trash2} onClick={() => onDelete(initial.id)}>Eliminar</Button>
          ) : <span />}
          <div style={{ display: "flex", gap: 8 }}>
            <Button variant="secondary" onClick={onClose}>Cancelar</Button>
            <Button
              variant="primary"
              icon={CheckCircle2}
              disabled={!canSave}
              onClick={() => {
                const accionFinal = usarPersonalizada ? accionPersonalizada.trim() : form.proximaAccion;
                if (usarPersonalizada && accionFinal && !acciones.includes(accionFinal)) onAddAccion(accionFinal);
                onSave({
                  ...form,
                  id: form.id || uid(),
                  proximaAccion: accionFinal,
                  valorEstimado: form.valorEstimado === "" ? 0 : Number(form.valorEstimado),
                });
              }}
            >
              Guardar oportunidad
            </Button>
          </div>
        </div>
      </Card>
    </div>
  );
}

function CRMCard({ op, onEdit, onDragStart }) {
  const vencida = op.fechaProximaAccion && op.fechaProximaAccion < nowLocalISO();
  const esHoy = op.fechaProximaAccion && op.fechaProximaAccion.slice(0, 10) === todayISO() && !vencida;
  const valor = Number(op.valorEstimado) || 0;
  const numSeg = (op.seguimientos || []).length;
  const segColor = numSeg === 0 ? MUTED : numSeg <= 2 ? ACCENT : numSeg <= 4 ? WARN : BAD;
  return (
    <div
      draggable
      onDragStart={(e) => { e.dataTransfer.effectAllowed = "move"; onDragStart(op.id); }}
      onClick={() => onEdit(op)}
      style={{
        background: "#fff", border: `1px solid ${LINE}`, borderLeft: `3px solid ${CRM_ETAPA_COLOR[op.etapa]}`,
        borderRadius: 9, padding: "9px 10px", cursor: "grab", boxShadow: "0 1px 2px rgba(14,42,71,0.05)",
        display: "flex", flexDirection: "column", gap: 5, position: "relative",
      }}
    >
      {numSeg > 0 && (
        <div
          title={`${numSeg} seguimiento(s) registrados`}
          style={{
            position: "absolute", top: -7, right: -7, width: 20, height: 20, borderRadius: "50%",
            background: segColor, color: "#fff", fontSize: 10.5, fontWeight: 800,
            display: "flex", alignItems: "center", justifyContent: "center", border: "2px solid #fff",
            boxShadow: "0 1px 2px rgba(14,42,71,0.15)",
          }}
        >
          {numSeg}
        </div>
      )}
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 6 }}>
        <div style={{ fontWeight: 700, fontSize: 12.5, color: INK }}>{op.cliente}</div>
        <Pill color={CRM_PRIORIDAD_COLOR[op.prioridad]} style={{ fontSize: 10, padding: "2px 7px", flexShrink: 0 }}>{op.prioridad}</Pill>
      </div>
      {(op.empresa || op.ciudad) && (
        <div style={{ fontSize: 11, color: MUTED, display: "flex", alignItems: "center", gap: 4 }}>
          {op.ciudad && <><MapPin size={11} />{op.ciudad}</>}{op.empresa ? ` · ${op.empresa}` : ""}
        </div>
      )}
      {op.telefono && (
        <div style={{ fontSize: 11, color: MUTED, display: "flex", alignItems: "center", gap: 4 }}>
          <Phone size={11} />{op.telefono}
        </div>
      )}
      {(valor > 0 || op.modelo) && (
        <div style={{ fontSize: 11.5, fontWeight: 700, color: INK, display: "flex", alignItems: "center", gap: 4 }}>
          <Wallet size={11} color={MUTED} />
          {op.modelo ? `${op.modelo}${valor > 0 ? " · " : ""}` : ""}{valor > 0 ? fmtMoney(valor) : ""}
        </div>
      )}
      <div style={{
        marginTop: 2, fontSize: 10.5, borderRadius: 7, padding: "4px 6px",
        background: vencida ? `${BAD}12` : esHoy ? `${WARN}12` : PAPER,
        color: vencida ? BAD : esHoy ? WARN : MUTED, fontWeight: vencida || esHoy ? 700 : 500,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
          <CalendarClock size={11} /> {fmtDateTime(op.fechaProximaAccion)}
        </div>
        <div style={{ marginTop: 1 }}>{op.proximaAccion}</div>
      </div>
    </div>
  );
}

function CRMKanbanBoard({ oportunidades, onStageChange, onEdit }) {
  const [dragId, setDragId] = useState(null);
  const [overCol, setOverCol] = useState(null);
  return (
    <div style={{ display: "flex", gap: 10, overflowX: "auto", paddingBottom: 6 }}>
      {CRM_ETAPAS.map((etapa) => {
        const items = oportunidades
          .filter((o) => o.etapa === etapa)
          .sort((a, b) => {
            const fa = a.fechaProximaAccion || "";
            const fb = b.fechaProximaAccion || "";
            if (!fa && !fb) return 0;
            if (!fa) return -1; // sin fecha programada = urgente, va arriba
            if (!fb) return 1;
            return fa < fb ? -1 : fa > fb ? 1 : 0;
          });
        return (
          <div
            key={etapa}
            onDragOver={(e) => { e.preventDefault(); setOverCol(etapa); }}
            onDragLeave={() => setOverCol((c) => (c === etapa ? null : c))}
            onDrop={(e) => {
              e.preventDefault();
              if (dragId) onStageChange(dragId, etapa);
              setDragId(null); setOverCol(null);
            }}
            style={{
              minWidth: 230, width: 230, flexShrink: 0, background: overCol === etapa ? `${CRM_ETAPA_COLOR[etapa]}0F` : PAPER,
              borderRadius: 12, border: `1px solid ${overCol === etapa ? CRM_ETAPA_COLOR[etapa] : LINE}`,
              display: "flex", flexDirection: "column", maxHeight: "68vh",
            }}
          >
            <div style={{ padding: "10px 12px", borderBottom: `2px solid ${CRM_ETAPA_COLOR[etapa]}` }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span style={{ width: 8, height: 8, borderRadius: 999, background: CRM_ETAPA_COLOR[etapa], flexShrink: 0 }} />
                <span style={{ fontWeight: 700, fontSize: 11.5, color: INK }}>{etapa}</span>
                <span style={{ marginLeft: "auto", fontSize: 11, color: MUTED, background: "#fff", borderRadius: 999, padding: "1px 7px", border: `1px solid ${LINE}` }}>{items.length}</span>
              </div>
              {items.some((o) => Number(o.valorEstimado) > 0) && (
                <div style={{ fontSize: 10.5, color: MUTED, marginTop: 3, fontWeight: 600 }}>
                  {fmtMoney(items.reduce((s, o) => s + (Number(o.valorEstimado) || 0), 0))} en juego
                </div>
              )}
            </div>
            <div style={{ padding: 8, display: "flex", flexDirection: "column", gap: 8, overflowY: "auto" }}>
              {items.map((op) => <CRMCard key={op.id} op={op} onEdit={onEdit} onDragStart={setDragId} />)}
              {!items.length && <div style={{ fontSize: 11, color: MUTED, textAlign: "center", padding: "16px 0" }}>Sin oportunidades</div>}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ============================== APP PRINCIPAL ============================== */

export default function App() {
  const [ventas, setVentas] = useState([]);
  const [catalogos, setCatalogos] = useState(DEFAULT_CATALOGOS);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("dashboard");
  const [modal, setModal] = useState(null); // { mode: 'new'|'edit', data }
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [quoteModal, setQuoteModal] = useState(null);
  const [monthFilter, setMonthFilter] = useState(new Date().toISOString().slice(0, 7));
  const [search, setSearch] = useState("");

  // --- Estado exclusivo del CRM (independiente de ventas/catalogos) ---
  const [oportunidades, setOportunidades] = useState([]);
  const [accionesCRM, setAccionesCRM] = useState(DEFAULT_ACCIONES_CRM);
  const [loadingCRM, setLoadingCRM] = useState(true);
  const [crmModal, setCrmModal] = useState(null); // { mode: 'new'|'edit', data }
  const [crmConfirmDelete, setCrmConfirmDelete] = useState(null);
  const [crmView, setCrmView] = useState("tablero"); // 'tablero' | 'panel' | 'recordatorios'
  const [searchCRM, setSearchCRM] = useState("");
  const [crmFiltroEtapa, setCrmFiltroEtapa] = useState("");
  const [crmFiltroPrioridad, setCrmFiltroPrioridad] = useState("");
  const [crmOrden, setCrmOrden] = useState("urgencia"); // 'urgencia' | 'valor' | 'cliente'
  const [searchClientes, setSearchClientes] = useState("");

  const [reportFilters, setReportFilters] = useState({
    mes: "", anio: "", cliente: "", modelo: "", estado: "", ciudad: "", tipoEnvio: "", etapa: "",
  });

  // --- Sesión de Supabase (nube) ---
  const [session, setSession] = useState(null);
  const [authReady, setAuthReady] = useState(false);
  const [offlineMode, setOfflineMode] = useState(false);
  const [syncError, setSyncError] = useState(null);
  const [syncOk, setSyncOk] = useState(false);

  useEffect(() => {
    onSyncError = (msg) => { setSyncError(msg); setSyncOk(false); };
    onSyncOk = () => setSyncOk(true);
    return () => { onSyncError = null; onSyncOk = null; };
  }, []);

  useEffect(() => {
    if (!syncOk) return;
    const t = setTimeout(() => setSyncOk(false), 2500);
    return () => clearTimeout(t);
  }, [syncOk]);

  const reloadVentasYCatalogos = useCallback(async () => {
    setLoading(true);
    const [data, cat] = await Promise.all([loadFromStorage(), loadCatalogos()]);
    if (data && Array.isArray(data)) {
      // Hay algo guardado (aunque esté vacío): respetarlo tal cual, sin sembrar datos de ejemplo.
      setVentas(data);
    } else if (currentSession) {
      // Sesión nueva en la nube sin nada aún: empezar vacío, no con datos de ejemplo.
      setVentas([]);
    } else {
      const seed = seedData();
      setVentas(seed);
      await saveToStorage(seed);
    }
    if (cat && typeof cat === "object") {
      setCatalogos({ ...DEFAULT_CATALOGOS, ...cat });
    } else {
      setCatalogos(DEFAULT_CATALOGOS);
      await saveCatalogos(DEFAULT_CATALOGOS);
    }
    setLoading(false);
  }, []);

  const reloadCRM = useCallback(async () => {
    setLoadingCRM(true);
    const [ops, acc] = await Promise.all([loadCRMOportunidades(), loadCRMAcciones()]);
    if (ops && Array.isArray(ops)) {
      setOportunidades(ops);
    } else if (currentSession) {
      setOportunidades([]);
    } else {
      const seed = crmSeedData();
      setOportunidades(seed);
      await saveCRMOportunidades(seed);
    }
    if (acc && Array.isArray(acc) && acc.length) {
      setAccionesCRM(acc);
    } else {
      setAccionesCRM(DEFAULT_ACCIONES_CRM);
      await saveCRMAcciones(DEFAULT_ACCIONES_CRM);
    }
    setLoadingCRM(false);
  }, []);

  useEffect(() => {
    (async () => {
      const cached = await loadCachedSession();
      if (cached) {
        currentSession = cached;
        setSession(cached);
      }
      setAuthReady(true);
    })();
  }, []);

  useEffect(() => {
    if (!authReady) return;
    reloadVentasYCatalogos();
    reloadCRM();
  }, [authReady, session]);

  const handleLogin = async (email, password) => {
    const s = await supaLogin(email, password);
    currentSession = s;
    setSession(s);
    await cacheSession(s);
  };

  const handleLogout = async () => {
    currentSession = null;
    setSession(null);
    await cacheSession(null);
  };

  const [syncingNow, setSyncingNow] = useState(false);
  const handleSyncNow = async () => {
    setSyncingNow(true);
    setSyncError(null);
    await Promise.all([
      saveToStorage(ventas),
      saveCatalogos(catalogos),
      saveCRMOportunidades(oportunidades),
      saveCRMAcciones(accionesCRM),
    ]);
    setSyncingNow(false);
  };

  const addCatalogItem = (key, value) => {
    setCatalogos((prev) => {
      const trimmed = value.trim();
      if (!trimmed || prev[key].includes(trimmed)) return prev;
      const next = { ...prev, [key]: [...prev[key], trimmed] };
      saveCatalogos(next);
      return next;
    });
  };

  const removeCatalogItem = (key, index) => {
    setCatalogos((prev) => {
      const next = { ...prev, [key]: prev[key].filter((_, i) => i !== index) };
      saveCatalogos(next);
      return next;
    });
  };

  const updateEmpresaInfo = (patch) => {
    setCatalogos((prev) => {
      const next = { ...prev, empresaInfo: { ...prev.empresaInfo, ...patch } };
      saveCatalogos(next);
      return next;
    });
  };

  // --- Carga y guardado exclusivos del CRM: ver reloadCRM(), disparado por el efecto de autenticación ---

  const crmSaveOpportunity = (record) => {
    setOportunidades((prev) => {
      const exists = prev.some((o) => o.id === record.id);
      const next = exists ? prev.map((o) => (o.id === record.id ? record : o)) : [record, ...prev];
      saveCRMOportunidades(next);
      return next;
    });
    setCrmModal(null);
  };

  const crmDeleteOpportunity = (id) => {
    setOportunidades((prev) => {
      const next = prev.filter((o) => o.id !== id);
      saveCRMOportunidades(next);
      return next;
    });
    setCrmConfirmDelete(null);
  };

  const crmChangeStage = (id, etapa) => {
    setOportunidades((prev) => {
      const next = prev.map((o) => (o.id === id ? { ...o, etapa } : o));
      saveCRMOportunidades(next);
      return next;
    });
  };

  const crmAddAccion = (nueva) => {
    setAccionesCRM((prev) => {
      if (prev.includes(nueva)) return prev;
      const next = [...prev, nueva];
      saveCRMAcciones(next);
      return next;
    });
  };

  /* ---------- Derivados del CRM ---------- */

  const crmActivas = useMemo(() => oportunidades.filter((o) => o.etapa !== "Inactivo / Perdido"), [oportunidades]);

  const oportunidadesFiltradas = useMemo(() => {
    if (!searchCRM.trim()) return oportunidades;
    const q = searchCRM.toLowerCase();
    return oportunidades.filter((o) =>
      [o.cliente, o.empresa, o.telefono, o.ciudad].some((f) => (f || "").toLowerCase().includes(q))
    );
  }, [oportunidades, searchCRM]);

  const crmContactarHoy = useMemo(
    () => crmActivas.filter((o) => o.fechaProximaAccion && o.fechaProximaAccion.slice(0, 10) <= todayISO()),
    [crmActivas]
  );
  const crmNegociacionesAbiertas = useMemo(() => oportunidades.filter((o) => o.etapa === "Negociación").length, [oportunidades]);
  const crmCotizacionesEsperando = useMemo(() => oportunidades.filter((o) => o.etapa === "Cotización enviada").length, [oportunidades]);
  const crmPorReactivar = useMemo(() => {
    const limite = new Date();
    limite.setDate(limite.getDate() - CRM_ENFRIANDOSE_DIAS);
    const limiteISO = limite.toISOString().slice(0, 10);
    return crmActivas.filter((o) => o.fechaUltimoContacto && o.fechaUltimoContacto <= limiteISO);
  }, [crmActivas]);

  const crmRecordatorios = useMemo(() => {
    const ahora = nowLocalISO();
    const hoyDate = todayISO();
    const atrasadas = crmActivas
      .filter((o) => o.fechaProximaAccion && o.fechaProximaAccion < ahora)
      .sort((a, b) => (a.fechaProximaAccion < b.fechaProximaAccion ? -1 : 1));
    const hoy = crmActivas
      .filter((o) => o.fechaProximaAccion && o.fechaProximaAccion.slice(0, 10) === hoyDate && o.fechaProximaAccion >= ahora)
      .sort((a, b) => (a.fechaProximaAccion < b.fechaProximaAccion ? -1 : 1));
    return { atrasadas, hoy };
  }, [crmActivas]);

  const crmPendientesPorAccion = useMemo(() => {
    const relevantes = [...crmRecordatorios.atrasadas, ...crmRecordatorios.hoy];
    const map = {};
    relevantes.forEach((o) => {
      const key = o.proximaAccion || "Sin acción definida";
      if (!map[key]) map[key] = [];
      map[key].push(o);
    });
    return Object.entries(map).sort((a, b) => b[1].length - a[1].length);
  }, [crmRecordatorios]);

  const crmResumenSemanal = useMemo(() => {
    const inicioSemana = new Date();
    const dia = inicioSemana.getDay();
    const diffToMonday = dia === 0 ? 6 : dia - 1;
    inicioSemana.setDate(inicioSemana.getDate() - diffToMonday);
    const inicioISO = `${inicioSemana.getFullYear()}-${String(inicioSemana.getMonth() + 1).padStart(2, "0")}-${String(inicioSemana.getDate()).padStart(2, "0")}`;
    let seguimientos = 0;
    oportunidades.forEach((o) => {
      (o.seguimientos || []).forEach((s) => { if (s.fecha && s.fecha.slice(0, 10) >= inicioISO) seguimientos++; });
    });
    const nuevas = oportunidades.filter((o) => o.creadoEn && o.creadoEn >= inicioISO).length;
    return { seguimientos, nuevas };
  }, [oportunidades]);

  const crmRecuperables = useMemo(() => {
    const limite = new Date();
    limite.setDate(limite.getDate() - 60);
    const limiteISO = limite.toISOString().slice(0, 10);
    return oportunidades.filter((o) => o.etapa === "Inactivo / Perdido" && (o.fechaUltimoContacto || o.creadoEn || "") <= limiteISO);
  }, [oportunidades]);

  const crmSaludo = useMemo(() => {
    const nombreSaludo = "José";
    const totalPendientes = crmRecordatorios.atrasadas.length + crmRecordatorios.hoy.length;
    if (!totalPendientes) return `Vas al día, ${nombreSaludo} — ningún pendiente en tu CRM por ahora. 🎉`;
    const siguiente = crmRecordatorios.atrasadas[0] || crmRecordatorios.hoy[0];
    const urgencia = crmRecordatorios.atrasadas.length ? "atrasado" : "para hoy";
    const extra = totalPendientes > 1 ? ` (y ${totalPendientes - 1} más pendiente${totalPendientes - 1 > 1 ? "s" : ""})` : "";
    return `Oye ${nombreSaludo}, lo más urgente es ${siguiente.cliente}: ${siguiente.proximaAccion} — ${urgencia}${siguiente.fechaProximaAccion ? ` (${fmtDateTime(siguiente.fechaProximaAccion)})` : ""}.${extra}`;
  }, [crmRecordatorios]);

  const persist = useCallback(async (next) => {
    setVentas(next);
    await saveToStorage(next);
  }, []);

  const handleSave = (record) => {
    setVentas((prev) => {
      const exists = prev.some((v) => v.id === record.id);
      const next = exists ? prev.map((v) => (v.id === record.id ? record : v)) : [record, ...prev];
      saveToStorage(next);
      return next;
    });
    setModal(null);
  };

  const handleDelete = (id) => {
    setVentas((prev) => {
      const next = prev.filter((v) => v.id !== id);
      saveToStorage(next);
      return next;
    });
    setConfirmDelete(null);
  };

  const handleStageChange = (id, etapa) => {
    setVentas((prev) => {
      const next = prev.map((v) => (v.id === id ? { ...v, etapa } : v));
      saveToStorage(next);
      return next;
    });
  };

  // Registra un pago (anticipo, pago parcial o finiquito) sobre un folio ya existente.
  const handleAddPago = (ventaId, pago) => {
    setVentas((prev) => {
      const next = prev.map((v) => {
        if (v.id !== ventaId) return v;
        const pagosActuales = getPagosPedido(v).filter((p) => !String(p.id).startsWith("legacy-"));
        const pagosNuevos = [...pagosActuales, pago];
        const saldoNuevo = round2(getSubtotalPedido(v) - pagosNuevos.reduce((s, p) => s + (Number(p.monto) || 0), 0));
        return {
          ...v,
          pagos: pagosNuevos,
          fechaRealCobro: pago.fecha || v.fechaRealCobro,
          etapa: saldoNuevo <= 0.5 ? "Cobrado" : v.etapa,
        };
      });
      saveToStorage(next);
      return next;
    });
  };

  /* ---------- Derivados generales ---------- */

  // Cada pago (anticipo/parcial/finiquito) de cada folio es un "movimiento de cobro"
  // independiente: cuenta para la comisión del mes de SU PROPIA fecha, sin duplicar el pedido.
  const movimientosCobro = useMemo(() => {
    const list = [];
    ventas.forEach((v) => {
      getPagosPedido(v).forEach((p) => {
        list.push({
          ...p, ventaId: v.id, numeroPedido: v.numeroPedido, cliente: v.cliente, empresa: v.empresa,
          modelo: v.modelo, estadoDestino: v.estadoDestino, ciudadDestino: v.ciudadDestino, conIVA: v.conIVA,
        });
      });
    });
    return list;
  }, [ventas]);

  const [y, mIdx] = monthFilter.split("-").map((x, i) => (i === 1 ? Number(x) - 1 : Number(x)));

  const movimientosMes = useMemo(
    () => movimientosCobro.filter((p) => monthKeyOf(p.fecha) === monthFilter),
    [movimientosCobro, monthFilter]
  );
  const pedidosMes = useMemo(
    () => ventas.filter((v) => monthKeyOf(v.fecha) === monthFilter),
    [ventas, monthFilter]
  );
  const pedidosConMovimientoMes = useMemo(() => {
    const ids = new Set(movimientosMes.map((p) => p.ventaId));
    return ventas.filter((v) => ids.has(v.id));
  }, [ventas, movimientosMes]);

  const subtotalMes = useMemo(() => movimientosMes.reduce((s, p) => s + (Number(p.monto) || 0), 0), [movimientosMes]);
  const totalConIVAMes = useMemo(
    () => movimientosMes.reduce((s, p) => s + (Number(p.monto) || 0) * (p.conIVA ? 1.16 : 1), 0),
    [movimientosMes]
  );
  const ivaMes = totalConIVAMes - subtotalMes;

  const tierActual = getTier(subtotalMes);
  const tierSiguiente = getNextTier(subtotalMes);
  const comisionActual = subtotalMes * (tierActual.pct / 100);
  const avancePct = Math.min(100, (subtotalMes / META_MENSUAL) * 100);
  const faltanteSiguiente = tierSiguiente ? tierSiguiente.min - subtotalMes : 0;

  const hoy = new Date();
  const esMesActual = monthFilter === hoy.toISOString().slice(0, 7);
  const diaDelMes = esMesActual ? hoy.getDate() : daysInMonth(y, mIdx);
  const totalDiasMes = daysInMonth(y, mIdx);
  const pronostico = diaDelMes > 0 ? (subtotalMes / diaDelMes) * totalDiasMes : subtotalMes;
  const tierPronostico = getTier(pronostico);
  const comisionEstimada = pronostico * (tierPronostico.pct / 100);

  const ticketPromedio = pedidosConMovimientoMes.length ? subtotalMes / pedidosConMovimientoMes.length : 0;
  const totalClientesMes = new Set(pedidosConMovimientoMes.map((v) => v.cliente)).size;
  const totalM2Mes = pedidosConMovimientoMes.reduce((s, v) => s + Number(v.metrosCuadrados || 0), 0);
  const totalPiezasMes = pedidosConMovimientoMes.reduce((s, v) => s + Number(v.cantidadPiezas || 0), 0);

  const mensajeNivel = subtotalMes < 250000
    ? `Te faltan ${fmtMoney(250000 - subtotalMes)} para comenzar a comisionar.`
    : `Actualmente estás en el nivel de comisión del ${tierActual.pct.toFixed(2)}%.`;

  /* ---------- Gráficas ---------- */

  const porDia = useMemo(() => {
    const map = {};
    for (let d = 1; d <= totalDiasMes; d++) map[d] = 0;
    movimientosMes.forEach((p) => {
      const d = Number((p.fecha || "").slice(8, 10));
      if (d) map[d] = (map[d] || 0) + (Number(p.monto) || 0);
    });
    return Object.keys(map).map((d) => ({ dia: d, ventas: Math.round(map[d]) }));
  }, [movimientosMes, totalDiasMes]);

  const porSemana = useMemo(() => {
    const map = {};
    movimientosMes.forEach((p) => {
      const d = Number((p.fecha || "").slice(8, 10));
      if (!d) return;
      const semana = `Sem ${Math.ceil(d / 7)}`;
      map[semana] = (map[semana] || 0) + (Number(p.monto) || 0);
    });
    return Object.keys(map).sort().map((k) => ({ semana: k, ventas: Math.round(map[k]) }));
  }, [movimientosMes]);

  const porMes = useMemo(() => {
    const map = {};
    movimientosCobro.forEach((p) => {
      const k = monthKeyOf(p.fecha);
      if (!k) return;
      map[k] = (map[k] || 0) + (Number(p.monto) || 0);
    });
    return Object.keys(map).sort().slice(-6).map((k) => {
      const [yy, mm] = k.split("-");
      return { mes: `${MESES[Number(mm) - 1]} ${yy.slice(2)}`, ventas: Math.round(map[k]) };
    });
  }, [movimientosCobro]);

  const historialComisiones = useMemo(() => {
    const map = {};
    movimientosCobro.forEach((p) => {
      const k = monthKeyOf(p.fecha);
      if (!k) return;
      map[k] = (map[k] || 0) + (Number(p.monto) || 0);
    });
    return Object.keys(map).sort().reverse().slice(0, 12).map((k) => {
      const [yy, mm] = k.split("-");
      const subtotalMesH = map[k];
      const tierH = getTier(subtotalMesH);
      return {
        key: k,
        mes: `${MESES[Number(mm) - 1]} ${yy}`,
        subtotal: subtotalMesH,
        pct: tierH.pct,
        comision: subtotalMesH * (tierH.pct / 100),
      };
    });
  }, [movimientosCobro]);

  const porEstado = useMemo(() => {
    const map = {};
    movimientosMes.forEach((p) => {
      map[p.estadoDestino] = (map[p.estadoDestino] || 0) + (Number(p.monto) || 0);
    });
    return Object.keys(map).map((k) => ({ name: k, value: Math.round(map[k]) })).sort((a, b) => b.value - a.value);
  }, [movimientosMes]);

  const porModelo = useMemo(() => {
    const map = {};
    movimientosMes.forEach((p) => {
      map[p.modelo] = (map[p.modelo] || 0) + (Number(p.monto) || 0);
    });
    return Object.keys(map).map((k) => ({ name: k, value: Math.round(map[k]) })).sort((a, b) => b.value - a.value);
  }, [movimientosMes]);

  const PIE_COLORS = [NAVY, ACCENT, "#7A4FC2", GOOD, WARN, "#2C8FAE", BAD, "#8C99AB"];

  /* ---------- Clientes ---------- */

  const clientesAgg = useMemo(() => {
    const map = {};
    ventas.forEach((v) => {
      const key = v.cliente || "—";
      if (!map[key]) map[key] = { cliente: key, empresa: v.empresa, pedidos: 0, totalCobrado: 0, ultima: v.fecha, telefono: v.telefono, correo: v.correo };
      map[key].pedidos += 1;
      if (v.fecha > map[key].ultima) map[key].ultima = v.fecha;
    });
    movimientosCobro.forEach((p) => {
      const key = p.cliente || "—";
      if (map[key]) map[key].totalCobrado += Number(p.monto) || 0;
    });
    return Object.values(map).sort((a, b) => b.totalCobrado - a.totalCobrado);
  }, [ventas, movimientosCobro]);

  /* ---------- Modelos ---------- */

  const modelosAgg = useMemo(() => {
    const map = {};
    ventas.forEach((v) => {
      const key = v.modelo || "—";
      if (!map[key]) map[key] = { modelo: key, piezas: 0, m2: 0, ventas: 0, clientes: new Set() };
      map[key].piezas += Number(v.cantidadPiezas || 0);
      map[key].m2 += Number(v.metrosCuadrados || 0);
      map[key].clientes.add(v.cliente);
    });
    movimientosCobro.forEach((p) => {
      const key = p.modelo || "—";
      if (map[key]) map[key].ventas += Number(p.monto) || 0;
    });
    return Object.values(map).map((m) => ({ ...m, clientes: m.clientes.size })).sort((a, b) => b.ventas - a.ventas);
  }, [ventas, movimientosCobro]);

  /* ---------- Alertas ---------- */

  const alertas = useMemo(() => {
    const list = [];
    if (tierSiguiente && faltanteSiguiente > 0 && faltanteSiguiente <= 30000) {
      list.push({ tipo: "nivel", icon: TrendingUp, color: ACCENT, texto: `Te faltan ${fmtMoney(faltanteSiguiente)} para subir al nivel de ${tierSiguiente.pct.toFixed(2)}%.` });
    }
    const recordatoriosHoy = ventas.filter((v) => v.proximoSeguimiento && v.proximoSeguimiento <= nowLocalISO() && !["Cobrado", "Entregado"].includes(v.etapa));
    if (recordatoriosHoy.length) {
      recordatoriosHoy.slice(0, 5).forEach((v) => {
        list.push({
          tipo: "recordatorio", icon: CalendarClock, color: ACCENT,
          texto: `${v.cliente}${v.notaSeguimiento ? ` — ${v.notaSeguimiento}` : " — Dar seguimiento"} (programado ${fmtDateTime(v.proximoSeguimiento)}, pedido ${v.numeroPedido}).`,
        });
      });
      if (recordatoriosHoy.length > 5) {
        list.push({ tipo: "recordatorio", icon: CalendarClock, color: ACCENT, texto: `+${recordatoriosHoy.length - 5} recordatorio(s) más pendientes.` });
      }
    }
    const pendientesCobro = ventas.filter((v) => getSubtotalPedido(v) > 0 && getSaldoPendiente(v) > 0.5);
    if (pendientesCobro.length) {
      list.push({ tipo: "cobro", icon: Wallet, color: WARN, texto: `${pendientesCobro.length} pedido(s) con saldo pendiente por cobrar.` });
    }
    const sinSeguimiento = ventas.filter((v) => {
      if (["Cobrado", "Entregado"].includes(v.etapa)) return false;
      const dias = (new Date() - new Date(v.fecha)) / 86400000;
      return dias > 10;
    });
    if (sinSeguimiento.length) {
      list.push({ tipo: "seguimiento", icon: Clock, color: BAD, texto: `${sinSeguimiento.length} cliente(s)/pedido(s) sin avanzar de etapa desde hace más de 10 días.` });
    }
    const ultimaVenta = ventas.reduce((max, v) => (v.fecha > max ? v.fecha : max), "");
    if (ultimaVenta) {
      const dias = Math.floor((new Date() - new Date(ultimaVenta)) / 86400000);
      if (dias >= 4) list.push({ tipo: "inactividad", icon: AlertTriangle, color: BAD, texto: `Llevas ${dias} días sin registrar una nueva venta.` });
    }
    if (subtotalMes >= META_MENSUAL && movimientosMes.length) {
      list.push({ tipo: "meta", icon: Sparkles, color: GOOD, texto: `¡Meta mensual alcanzada! Ya vas en ${tierActual.pct.toFixed(2)}% de comisión este mes.` });
    }
    return list;
  }, [ventas, tierSiguiente, faltanteSiguiente, subtotalMes, movimientosMes, tierActual]);

  /* ---------- Exportaciones ---------- */

  const buildReportRows = () => {
    let rows = ventas.filter((v) => {
      if (reportFilters.mes && v.fecha.slice(5, 7) !== reportFilters.mes) return false;
      if (reportFilters.anio && v.fecha.slice(0, 4) !== reportFilters.anio) return false;
      if (reportFilters.cliente && !v.cliente.toLowerCase().includes(reportFilters.cliente.toLowerCase())) return false;
      if (reportFilters.modelo && v.modelo !== reportFilters.modelo) return false;
      if (reportFilters.estado && v.estadoDestino !== reportFilters.estado) return false;
      if (reportFilters.ciudad && !v.ciudadDestino.toLowerCase().includes(reportFilters.ciudad.toLowerCase())) return false;
      if (reportFilters.tipoEnvio && v.tipoEnvio !== reportFilters.tipoEnvio) return false;
      if (reportFilters.etapa && v.etapa !== reportFilters.etapa) return false;
      return true;
    });
    return rows.sort((a, b) => (a.fecha < b.fecha ? 1 : -1));
  };

  const reportRows = useMemo(buildReportRows, [ventas, reportFilters]);

  const toExportArray = (rows) => rows.map((v) => {
    const subtotal = getSubtotalPedido(v);
    const totalPagado = getTotalPagado(v);
    const saldo = getSaldoPendiente(v);
    return {
      Fecha: v.fecha, Pedido: v.numeroPedido, Cliente: v.cliente, Empresa: v.empresa,
      Teléfono: v.telefono, Correo: v.correo, Modelo: v.modelo, Piezas: v.cantidadPiezas,
      "M2": v.metrosCuadrados, Color: v.color,
      Acabado: v.acabado, "Estado destino": v.estadoDestino, "Ciudad destino": v.ciudadDestino,
      "Tipo envío": v.tipoEnvio, Fletera: v.fletera, "Con IVA": v.conIVA ? "Sí" : "No",
      "Subtotal sin IVA": round2(subtotal), "Total con IVA": v.conIVA ? round2(subtotal * 1.16) : "",
      "Forma de pago": v.formaPago, "% Anticipo sugerido": v.anticipoPct || "",
      "Total pagado": round2(totalPagado), "Saldo pendiente": round2(saldo),
      Etapa: getLiquidado(v) ? "Liquidado" : v.etapa, Observaciones: v.observaciones,
    };
  });

  const exportCSV = () => {
    const rows = toExportArray(reportRows);
    if (!rows.length) return;
    const headers = Object.keys(rows[0]);
    const csv = [headers.join(","), ...rows.map((r) => headers.map((h) => `"${String(r[h] ?? "").replace(/"/g, '""')}"`).join(","))].join("\n");
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `reporte-ventas-${todayISO()}.csv`; a.click();
    URL.revokeObjectURL(url);
  };

  const exportExcel = () => {
    const rows = toExportArray(reportRows);
    if (!rows.length) return;
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Ventas");
    XLSX.writeFile(wb, `reporte-ventas-${todayISO()}.xlsx`);
  };

  const exportPDF = () => {
    printClean("Reporte de ventas");
  };

  /* ---------- Búsqueda tabla ventas ---------- */

  const movimientosOrdenados = useMemo(() => {
    const base = [...movimientosCobro].sort((a, b) => (a.fecha < b.fecha ? 1 : a.fecha > b.fecha ? -1 : 0));
    if (!search.trim()) return base;
    const q = search.toLowerCase();
    return base.filter((m) =>
      [m.cliente, m.empresa, m.numeroPedido, m.modelo, m.ciudadDestino, m.estadoDestino]
        .some((f) => (f || "").toLowerCase().includes(q))
    );
  }, [movimientosCobro, search]);

  /* ============================== NAV ============================== */

  const NAV = [
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "crm", label: "CRM", icon: Users2, badge: crmContactarHoy.length },
    { id: "ventas", label: "Ventas", icon: ClipboardList },
    { id: "clientes", label: "Clientes", icon: Users },
    { id: "modelos", label: "Modelos", icon: Boxes },
    { id: "reportes", label: "Reportes", icon: FileBarChart },
    { id: "alertas", label: "Alertas", icon: BellRing, badge: alertas.length },
    { id: "catalogo", label: "Catálogo", icon: Layers },
  ];

  if (!authReady) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: 400, color: MUTED, fontFamily: "inherit" }}>
        Verificando sesión…
      </div>
    );
  }

  if (!session && !offlineMode) {
    return <LoginScreen onLogin={handleLogin} onOfflineMode={() => setOfflineMode(true)} />;
  }

  if (loading) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: 400, color: MUTED, fontFamily: "inherit" }}>
        Cargando panel comercial…
      </div>
    );
  }

  return (
    <div style={{
      fontFamily: "'Segoe UI', 'Inter', system-ui, -apple-system, sans-serif",
      background: PAPER, color: INK, minHeight: 640, borderRadius: 16, overflow: "hidden",
      border: `1px solid ${LINE}`, display: "flex", flexDirection: "column",
    }}>
      <style>{`
        @media print {
          .no-print { display: none !important; }
          .print-area { box-shadow: none !important; border: none !important; }
        }
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: #C7D2DE; border-radius: 8px; }
        .spin { animation: spin-rotate 0.8s linear infinite; }
        @keyframes spin-rotate { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
      `}</style>

      {/* HEADER */}
      <div style={{ position: "relative", background: `linear-gradient(120deg, ${NAVY} 0%, ${NAVY_LIGHT} 100%)`, padding: "18px 24px", overflow: "hidden" }} className="no-print-bg">
        <CelosiaBand height={90} />
        <div style={{ position: "relative", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: "rgba(255,255,255,0.12)", display: "flex", alignItems: "center", justifyContent: "center", border: "1px solid rgba(255,255,255,0.25)" }}>
              <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
                <rect x="1" y="1" width="8" height="8" rx="1.5" stroke="#fff" strokeWidth="1.6" />
                <rect x="11" y="1" width="8" height="8" rx="1.5" stroke="#fff" strokeWidth="1.6" />
                <rect x="1" y="11" width="8" height="8" rx="1.5" stroke="#fff" strokeWidth="1.6" />
                <rect x="11" y="11" width="8" height="8" rx="1.5" stroke="#fff" strokeWidth="1.6" />
              </svg>
            </div>
            <div>
              <div style={{ color: "#fff", fontWeight: 800, fontSize: 17, letterSpacing: 0.2 }}>Panel Comercial · Celosías</div>
              <div style={{ color: "rgba(255,255,255,0.75)", fontSize: 12.5 }}>Ventas, cobranza y comisiones</div>
            </div>
          </div>
          <div className="no-print" style={{ display: "flex", alignItems: "center", gap: 10 }}>
            {session ? (
              <>
                <button
                  onClick={handleSyncNow}
                  disabled={syncingNow}
                  title="Forzar sincronización con la nube ahora"
                  style={{
                    border: "1px solid rgba(255,255,255,0.3)", background: "rgba(255,255,255,0.1)", color: "#fff",
                    borderRadius: 8, padding: "5px 9px", fontSize: 11.5, cursor: syncingNow ? "default" : "pointer",
                    display: "flex", alignItems: "center", gap: 5, opacity: syncingNow ? 0.6 : 1,
                  }}
                >
                  <RefreshCw size={12} className={syncingNow ? "spin" : ""} />
                  {syncingNow ? "Sincronizando…" : "Sincronizar ahora"}
                </button>
                <span
                  onClick={handleLogout}
                  title="Cerrar sesión"
                  style={{ color: "rgba(255,255,255,0.8)", fontSize: 11.5, cursor: "pointer", display: "flex", alignItems: "center", gap: 5 }}
                >
                  ☁ {session.email} · salir
                </span>
              </>
            ) : (
              <span style={{ color: "rgba(255,255,255,0.6)", fontSize: 11.5 }}>Sin conexión (solo este chat)</span>
            )}
            <input
              type="month" value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}
              style={{ border: "1px solid rgba(255,255,255,0.3)", background: "rgba(255,255,255,0.1)", color: "#fff", borderRadius: 9, padding: "7px 10px", fontSize: 13 }}
            />
            <Button icon={Plus} onClick={() => setModal({ mode: "new" })} style={{ background: "#fff", color: NAVY }}>
              Nuevo pedido
            </Button>
          </div>
        </div>

        {/* NAV TABS */}
        <div className="no-print" style={{ position: "relative", display: "flex", gap: 4, marginTop: 16, flexWrap: "wrap" }}>
          {NAV.map((n) => {
            const active = tab === n.id;
            return (
              <button
                key={n.id}
                onClick={() => setTab(n.id)}
                style={{
                  display: "flex", alignItems: "center", gap: 6, padding: "8px 13px", borderRadius: 9,
                  border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600,
                  background: active ? "#fff" : "rgba(255,255,255,0.08)",
                  color: active ? NAVY : "rgba(255,255,255,0.85)",
                }}
              >
                <n.icon size={15} />
                {n.label}
                {!!n.badge && (
                  <span style={{ background: active ? BAD : "#fff", color: active ? "#fff" : BAD, borderRadius: 999, fontSize: 10.5, fontWeight: 800, padding: "1px 6px", marginLeft: 2 }}>
                    {n.badge}
                  </span>
                )}
              </button>
            );
          })}
        </div>
      </div>

      {syncError && (
        <div className="no-print" style={{ background: `${BAD}12`, borderBottom: `1px solid ${BAD}33`, padding: "8px 20px", display: "flex", alignItems: "center", gap: 10, fontSize: 12.5, color: BAD }}>
          <AlertTriangle size={15} />
          <span style={{ flex: 1 }}>No se pudo sincronizar con la nube: {syncError}. Tus cambios se guardaron aquí en este dispositivo, pero no en Supabase todavía.</span>
          <button onClick={() => setSyncError(null)} style={{ border: "none", background: "transparent", cursor: "pointer", color: BAD }}><X size={15} /></button>
        </div>
      )}
      {syncOk && !syncError && (
        <div className="no-print" style={{ background: `${GOOD}12`, borderBottom: `1px solid ${GOOD}33`, padding: "6px 20px", display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: GOOD }}>
          <CheckCircle2 size={13} /> Sincronizado con la nube.
        </div>
      )}

      {/* BODY */}
      <div style={{ padding: 20, flex: 1, overflowY: "auto" }} className="print-area">

        {tab === "dashboard" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>

            {/* Barra de progreso de meta */}
            <Card style={{ padding: 18 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 10 }}>
                <div>
                  <div style={{ fontSize: 12.5, color: MUTED, fontWeight: 600 }}>META MENSUAL: {fmtMoney(META_MENSUAL)} sin IVA</div>
                  <div style={{ fontSize: 20, fontWeight: 800, color: INK }}>{fmtMoney(subtotalMes)} <span style={{ fontWeight: 500, fontSize: 13, color: MUTED }}>acumulado sin IVA</span></div>
                </div>
                <Pill color={subtotalMes >= META_MENSUAL ? GOOD : ACCENT}>{mensajeNivel}</Pill>
              </div>
              <div style={{ height: 10, borderRadius: 999, background: LINE, overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${avancePct}%`, background: `linear-gradient(90deg, ${NAVY}, ${ACCENT})`, borderRadius: 999, transition: "width .4s" }} />
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6, fontSize: 12, color: MUTED }}>
                <span>{avancePct.toFixed(1)}% de la meta mínima</span>
                {tierSiguiente ? <span>Faltan {fmtMoney(faltanteSiguiente)} para {tierSiguiente.pct.toFixed(2)}%</span> : <span>Nivel máximo alcanzado (3.50%)</span>}
              </div>
            </Card>

            {/* Stat cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: 12 }}>
              <StatCard icon={Wallet} label="Venta acum. sin IVA" value={fmtMoney(subtotalMes)} accent={NAVY} />
              <StatCard icon={Receipt} label="Venta acum. con IVA" value={fmtMoney(totalConIVAMes)} accent={ACCENT} />
              <StatCard icon={Receipt} label="IVA acumulado" value={fmtMoney(ivaMes)} accent={WARN} />
              <StatCard icon={ShoppingCart} label="Total de pedidos" value={fmtNum(pedidosMes.length)} sub={`${pedidosConMovimientoMes.length} con cobro`} accent={NAVY_SOFT} />
              <StatCard icon={Users2} label="Total de clientes" value={fmtNum(totalClientesMes)} accent={"#7A4FC2"} />
              <StatCard icon={Ruler} label="m² vendidos" value={fmtNum(totalM2Mes)} accent={"#2C8FAE"} />
              <StatCard icon={Package} label="Piezas vendidas" value={fmtNum(totalPiezasMes)} accent={"#2C8F5B"} />
              <StatCard icon={TrendingUp} label="Comisión actual" value={fmtMoney(comisionActual)} sub={`${tierActual.pct.toFixed(2)}%`} accent={GOOD} />
              <StatCard icon={Target} label="Comisión estimada" value={fmtMoney(comisionEstimada)} sub="al cierre del mes" accent={GOOD} />
              <StatCard icon={FileBarChart} label="Ticket promedio" value={fmtMoney(ticketPromedio)} accent={NAVY} />
              <StatCard icon={CalendarClock} label="Pronóstico de cierre" value={fmtMoney(pronostico)} sub={`día ${diaDelMes} de ${totalDiasMes}`} accent={ACCENT} />
              <StatCard icon={Target} label="Avance de meta" value={`${avancePct.toFixed(1)}%`} accent={WARN} />
            </div>

            {/* Charts */}
            <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 12 }}>
              <Card style={{ padding: 16 }}>
                <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 10 }}>Ventas por día (sin IVA)</div>
                <ResponsiveContainer width="100%" height={220}>
                  <AreaChart data={porDia}>
                    <defs>
                      <linearGradient id="gradVentas" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor={NAVY} stopOpacity={0.35} />
                        <stop offset="100%" stopColor={NAVY} stopOpacity={0} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={LINE} vertical={false} />
                    <XAxis dataKey="dia" tick={{ fontSize: 11, fill: MUTED }} axisLine={false} tickLine={false} />
                    <YAxis tick={{ fontSize: 11, fill: MUTED }} axisLine={false} tickLine={false} tickFormatter={(v) => `${v / 1000}k`} />
                    <Tooltip formatter={(v) => fmtMoney(v)} labelFormatter={(l) => `Día ${l}`} />
                    <Area type="monotone" dataKey="ventas" stroke={NAVY} fill="url(#gradVentas)" strokeWidth={2} />
                  </AreaChart>
                </ResponsiveContainer>
              </Card>

              <Card style={{ padding: 16 }}>
                <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 10 }}>Ventas por semana</div>
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={porSemana}>
                    <CartesianGrid strokeDasharray="3 3" stroke={LINE} vertical={false} />
                    <XAxis dataKey="semana" tick={{ fontSize: 11, fill: MUTED }} axisLine={false} tickLine={false} />
                    <YAxis tick={{ fontSize: 11, fill: MUTED }} axisLine={false} tickLine={false} tickFormatter={(v) => `${v / 1000}k`} />
                    <Tooltip formatter={(v) => fmtMoney(v)} />
                    <Bar dataKey="ventas" fill={ACCENT} radius={[6, 6, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </Card>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
              <Card style={{ padding: 16 }}>
                <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 10 }}>Últimos 6 meses</div>
                <ResponsiveContainer width="100%" height={200}>
                  <LineChart data={porMes}>
                    <CartesianGrid strokeDasharray="3 3" stroke={LINE} vertical={false} />
                    <XAxis dataKey="mes" tick={{ fontSize: 11, fill: MUTED }} axisLine={false} tickLine={false} />
                    <YAxis tick={{ fontSize: 11, fill: MUTED }} axisLine={false} tickLine={false} tickFormatter={(v) => `${v / 1000}k`} />
                    <Tooltip formatter={(v) => fmtMoney(v)} />
                    <Line type="monotone" dataKey="ventas" stroke={NAVY} strokeWidth={2.5} dot={{ r: 3 }} />
                  </LineChart>
                </ResponsiveContainer>
              </Card>

              <Card style={{ padding: 16 }}>
                <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 10 }}>Por estado destino</div>
                <ResponsiveContainer width="100%" height={200}>
                  <PieChart>
                    <Pie data={porEstado} dataKey="value" nameKey="name" innerRadius={45} outerRadius={75} paddingAngle={2}>
                      {porEstado.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                    </Pie>
                    <Tooltip formatter={(v) => fmtMoney(v)} />
                  </PieChart>
                </ResponsiveContainer>
              </Card>

              <Card style={{ padding: 16 }}>
                <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 10 }}>Por modelo</div>
                <ResponsiveContainer width="100%" height={200}>
                  <PieChart>
                    <Pie data={porModelo} dataKey="value" nameKey="name" innerRadius={45} outerRadius={75} paddingAngle={2}>
                      {porModelo.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                    </Pie>
                    <Tooltip formatter={(v) => fmtMoney(v)} />
                  </PieChart>
                </ResponsiveContainer>
              </Card>
            </div>

            <Card style={{ padding: 16 }}>
              <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 10 }}>Historial de comisiones (últimos 12 meses cerrados)</div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr style={{ textAlign: "left" }}>
                      {["Mes", "Venta sin IVA", "Nivel", "Comisión"].map((h) => (
                        <th key={h} style={{ padding: "8px 10px", fontSize: 11, color: MUTED, fontWeight: 700, textTransform: "uppercase" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {historialComisiones.map((h) => (
                      <tr key={h.key} style={{ borderTop: `1px solid ${LINE}` }}>
                        <td style={{ padding: "8px 10px", fontWeight: 600 }}>{h.mes}</td>
                        <td style={{ padding: "8px 10px" }}>{fmtMoney(h.subtotal)}</td>
                        <td style={{ padding: "8px 10px" }}><Pill color={h.pct > 0 ? GOOD : MUTED}>{h.pct.toFixed(2)}%</Pill></td>
                        <td style={{ padding: "8px 10px", fontWeight: 700 }}>{fmtMoney(h.comision)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {!historialComisiones.length && <EmptyState icon={FileBarChart} title="Aún no hay meses cerrados" sub="Cuando marques pedidos como Cobrado, aquí verás tu historial mes a mes." />}
              </div>
            </Card>
          </div>
        )}

        {tab === "crm" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Card style={{
              padding: "14px 16px", display: "flex", alignItems: "center", gap: 12,
              background: `linear-gradient(90deg, ${NAVY}0D, ${ACCENT}0D)`, border: `1px solid ${NAVY}22`,
            }}>
              <div style={{ width: 34, height: 34, borderRadius: 999, background: NAVY, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <Sparkles size={16} color="#fff" />
              </div>
              <div style={{ fontSize: 13, color: INK, fontWeight: 600 }}>{crmSaludo}</div>
            </Card>

            <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                <div style={{ display: "flex", background: PAPER, borderRadius: 9, padding: 3, border: `1px solid ${LINE}` }}>
                  {[
                    { id: "tablero", label: "Tablero" },
                    { id: "lista", label: "Lista" },
                    { id: "panel", label: "Panel" },
                    { id: "recordatorios", label: "Recordatorios", badge: crmRecordatorios.atrasadas.length + crmRecordatorios.hoy.length },
                  ].map((v) => (
                    <button
                      key={v.id}
                      onClick={() => setCrmView(v.id)}
                      style={{
                        border: "none", cursor: "pointer", borderRadius: 7, padding: "6px 12px", fontSize: 12.5, fontWeight: 600,
                        background: crmView === v.id ? "#fff" : "transparent", color: crmView === v.id ? NAVY : MUTED,
                        boxShadow: crmView === v.id ? "0 1px 2px rgba(14,42,71,0.1)" : "none", display: "flex", alignItems: "center", gap: 6,
                      }}
                    >
                      {v.label}
                      {!!v.badge && (
                        <span style={{ background: BAD, color: "#fff", borderRadius: 999, fontSize: 10, fontWeight: 800, padding: "1px 6px" }}>{v.badge}</span>
                      )}
                    </button>
                  ))}
                </div>
                <div style={{ position: "relative", width: 240 }}>
                  <Search size={14} style={{ position: "absolute", left: 9, top: 9, color: MUTED }} />
                  <TextInput
                    placeholder="Buscar contacto, empresa, tel, ciudad…"
                    value={searchCRM}
                    onChange={(e) => setSearchCRM(e.target.value)}
                    style={{ paddingLeft: 30, fontSize: 12.5 }}
                  />
                </div>
              </div>
              <Button icon={Plus} onClick={() => setCrmModal({ mode: "new" })}>Nueva oportunidad</Button>
            </div>

            {searchCRM.trim() && (
              <div style={{ fontSize: 12, color: MUTED }}>
                {oportunidadesFiltradas.length
                  ? `${oportunidadesFiltradas.length} contacto(s) encontrado(s) para "${searchCRM}".`
                  : `No hay ningún contacto registrado con "${searchCRM}" — puedes darlo de alta con "Nueva oportunidad".`}
              </div>
            )}

            {loadingCRM ? (
              <div style={{ padding: 40, textAlign: "center", color: MUTED, fontSize: 13 }}>Cargando CRM…</div>
            ) : crmView === "tablero" ? (
              <>
                <div style={{ fontSize: 12, color: MUTED, display: "flex", alignItems: "center", gap: 5 }}>
                  <Layers size={13} /> Cada columna se ordena sola: arriba lo urgente (sin fecha o vencido), abajo lo ya programado. Arrastra una tarjeta para cambiarla de etapa, o haz clic para editarla.
                </div>
                <CRMKanbanBoard oportunidades={oportunidadesFiltradas} onStageChange={crmChangeStage} onEdit={(o) => setCrmModal({ mode: "edit", data: o })} />
              </>
            ) : crmView === "lista" ? (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  <Select value={crmFiltroEtapa} onChange={(e) => setCrmFiltroEtapa(e.target.value)} style={{ width: 190 }}>
                    <option value="">Todas las etapas</option>
                    {CRM_ETAPAS.map((e) => <option key={e} value={e}>{e}</option>)}
                  </Select>
                  <Select value={crmFiltroPrioridad} onChange={(e) => setCrmFiltroPrioridad(e.target.value)} style={{ width: 160 }}>
                    <option value="">Todas las prioridades</option>
                    {CRM_PRIORIDADES.map((p) => <option key={p} value={p}>{p}</option>)}
                  </Select>
                  <Select value={crmOrden} onChange={(e) => setCrmOrden(e.target.value)} style={{ width: 200 }}>
                    <option value="urgencia">Ordenar: más urgente primero</option>
                    <option value="valor">Ordenar: mayor valor primero</option>
                    <option value="cliente">Ordenar: cliente (A-Z)</option>
                  </Select>
                  {(crmFiltroEtapa || crmFiltroPrioridad) && (
                    <Button size="sm" variant="ghost" onClick={() => { setCrmFiltroEtapa(""); setCrmFiltroPrioridad(""); }}>Quitar filtros</Button>
                  )}
                </div>
              <Card style={{ overflow: "hidden" }}>
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                    <thead>
                      <tr style={{ background: PAPER, textAlign: "left" }}>
                        {["Cliente", "Empresa / Ciudad", "Teléfono", "Modelo", "Valor", "Etapa", "Próxima acción", "Prioridad", "Seg."].map((h) => (
                          <th key={h} style={{ padding: "9px 10px", fontSize: 10.5, color: MUTED, fontWeight: 700, textTransform: "uppercase", whiteSpace: "nowrap" }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {oportunidadesFiltradas
                        .filter((o) => !crmFiltroEtapa || o.etapa === crmFiltroEtapa)
                        .filter((o) => !crmFiltroPrioridad || o.prioridad === crmFiltroPrioridad)
                        .sort((a, b) => {
                          if (crmOrden === "valor") return (Number(b.valorEstimado) || 0) - (Number(a.valorEstimado) || 0);
                          if (crmOrden === "cliente") return (a.cliente || "").localeCompare(b.cliente || "");
                          const fa = a.fechaProximaAccion || "";
                          const fb = b.fechaProximaAccion || "";
                          if (!fa && !fb) return 0;
                          if (!fa) return -1;
                          if (!fb) return 1;
                          return fa < fb ? -1 : fa > fb ? 1 : 0;
                        })
                        .map((o) => (
                        <tr
                          key={o.id}
                          onClick={() => setCrmModal({ mode: "edit", data: o })}
                          style={{ borderTop: `1px solid ${LINE}`, cursor: "pointer" }}
                        >
                          <td style={{ padding: "8px 10px", fontWeight: 600 }}>{o.cliente}</td>
                          <td style={{ padding: "8px 10px", color: MUTED }}>{[o.empresa, o.ciudad].filter(Boolean).join(" · ") || "—"}</td>
                          <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>{o.telefono || "—"}</td>
                          <td style={{ padding: "8px 10px" }}>{o.modelo || "—"}</td>
                          <td style={{ padding: "8px 10px", fontWeight: 600, whiteSpace: "nowrap" }}>{Number(o.valorEstimado) > 0 ? fmtMoney(o.valorEstimado) : "—"}</td>
                          <td style={{ padding: "8px 10px" }}><Pill color={CRM_ETAPA_COLOR[o.etapa]}>{o.etapa}</Pill></td>
                          <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>
                            <div>{o.proximaAccion}</div>
                            <div style={{ fontSize: 10.5, color: (o.fechaProximaAccion || "") < nowLocalISO() ? BAD : MUTED }}>{fmtDateTime(o.fechaProximaAccion)}</div>
                          </td>
                          <td style={{ padding: "8px 10px" }}><Pill color={CRM_PRIORIDAD_COLOR[o.prioridad]}>{o.prioridad}</Pill></td>
                          <td style={{ padding: "8px 10px", textAlign: "center", fontWeight: 700, color: MUTED }}>{(o.seguimientos || []).length}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {!oportunidadesFiltradas.length && <EmptyState icon={Users2} title="Sin contactos" sub="Registra tu primera oportunidad con el botón “Nueva oportunidad”." />}
                </div>
              </Card>
              </div>
            ) : crmView === "panel" ? (
              <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: 12 }}>
                  <StatCard icon={CalendarClock} label="Contactar hoy" value={fmtNum(crmContactarHoy.length)} sub="vencidas o programadas para hoy" accent={crmContactarHoy.length ? BAD : GOOD} />
                  <StatCard icon={Sparkles} label="Negociaciones abiertas" value={fmtNum(crmNegociacionesAbiertas)} accent={"#7A4FC2"} />
                  <StatCard icon={FileBarChart} label="Cotizaciones esperando respuesta" value={fmtNum(crmCotizacionesEsperando)} accent={ACCENT} />
                  <StatCard icon={Clock} label="Clientes por reactivar" value={fmtNum(crmPorReactivar.length)} sub={`sin contacto ${CRM_ENFRIANDOSE_DIAS}+ días`} accent={WARN} />
                </div>
                <Card style={{ padding: 16 }}>
                  <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 10 }}>Oportunidades por etapa</div>
                  <div style={{ display: "grid", gridTemplateColumns: `repeat(${CRM_ETAPAS.length}, 1fr)`, gap: 8 }}>
                    {CRM_ETAPAS.map((e) => {
                      const count = oportunidades.filter((o) => o.etapa === e).length;
                      return (
                        <div key={e} style={{ textAlign: "center", padding: "10px 6px", borderRadius: 10, background: PAPER, border: `1px solid ${LINE}` }}>
                          <div style={{ width: 8, height: 8, borderRadius: 999, background: CRM_ETAPA_COLOR[e], margin: "0 auto 6px" }} />
                          <div style={{ fontSize: 10.5, color: MUTED, fontWeight: 600, lineHeight: 1.2 }}>{e}</div>
                          <div style={{ fontSize: 16, fontWeight: 800, color: INK }}>{count}</div>
                        </div>
                      );
                    })}
                  </div>
                </Card>

                <Card style={{ padding: 16 }}>
                  <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 10 }}>Esta semana</div>
                  <div style={{ display: "flex", gap: 24, flexWrap: "wrap" }}>
                    <div>
                      <div style={{ fontSize: 22, fontWeight: 800, color: NAVY }}>{crmResumenSemanal.seguimientos}</div>
                      <div style={{ fontSize: 11.5, color: MUTED }}>seguimientos dados</div>
                    </div>
                    <div>
                      <div style={{ fontSize: 22, fontWeight: 800, color: NAVY }}>{crmResumenSemanal.nuevas}</div>
                      <div style={{ fontSize: 11.5, color: MUTED }}>oportunidades nuevas</div>
                    </div>
                  </div>
                </Card>

                {crmRecuperables.length > 0 && (
                  <Card style={{ padding: 16 }}>
                    <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 4 }}>Vale la pena reintentar</div>
                    <div style={{ fontSize: 12, color: MUTED, marginBottom: 10 }}>Perdidos hace 60+ días — a veces el que dijo que no, después sí compra.</div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                      {crmRecuperables.map((o) => (
                        <div
                          key={o.id}
                          onClick={() => setCrmModal({ mode: "edit", data: o })}
                          style={{ fontSize: 12.5, display: "flex", justifyContent: "space-between", gap: 8, cursor: "pointer", padding: "6px 8px", borderRadius: 7, background: PAPER }}
                        >
                          <span><b>{o.cliente}</b>{o.motivoPerdida ? ` — ${o.motivoPerdida}` : ""}</span>
                          <span style={{ color: MUTED, flexShrink: 0 }}>{fmtDate(o.fechaUltimoContacto || o.creadoEn)}</span>
                        </div>
                      ))}
                    </div>
                  </Card>
                )}
              </div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 8 }}>Pendientes por tipo de acción</div>
                  {crmPendientesPorAccion.length ? (
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 10 }}>
                      {crmPendientesPorAccion.map(([accion, items]) => (
                        <Card key={accion} style={{ padding: 12 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                            <div style={{ fontWeight: 700, fontSize: 12.5, color: INK }}>{accion}</div>
                            <span style={{ marginLeft: "auto", background: NAVY, color: "#fff", borderRadius: 999, fontSize: 10.5, fontWeight: 800, padding: "1px 7px" }}>{items.length}</span>
                          </div>
                          <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                            {items.map((o) => {
                              const atrasada = o.fechaProximaAccion < nowLocalISO() && o.fechaProximaAccion.slice(0, 10) !== todayISO();
                              return (
                                <div
                                  key={o.id}
                                  onClick={() => setCrmModal({ mode: "edit", data: o })}
                                  style={{ fontSize: 12, display: "flex", justifyContent: "space-between", gap: 6, cursor: "pointer", padding: "4px 6px", borderRadius: 6 }}
                                  onMouseEnter={(e) => (e.currentTarget.style.background = PAPER)}
                                  onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                                >
                                  <span>{o.cliente}</span>
                                  <span style={{ color: atrasada ? BAD : MUTED, fontWeight: atrasada ? 700 : 500, flexShrink: 0 }}>
                                    {o.fechaProximaAccion.slice(0, 10) === todayISO() ? "hoy" : fmtDate(o.fechaProximaAccion)}
                                  </span>
                                </div>
                              );
                            })}
                          </div>
                        </Card>
                      ))}
                    </div>
                  ) : (
                    <div style={{ fontSize: 12.5, color: MUTED }}>Nada pendiente por tipo de acción ahora mismo.</div>
                  )}
                </div>

                <div>
                  <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 8, color: BAD }}>Atrasadas ({crmRecordatorios.atrasadas.length})</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    {crmRecordatorios.atrasadas.map((o) => (
                      <Card key={o.id} style={{ padding: 12, display: "flex", alignItems: "center", gap: 10, cursor: "pointer", borderColor: `${BAD}44` }} onClick={() => setCrmModal({ mode: "edit", data: o })}>
                        <div style={{ width: 32, height: 32, borderRadius: 8, background: `${BAD}16`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                          <AlertTriangle size={16} color={BAD} />
                        </div>
                        <div style={{ fontSize: 13 }}>
                          <b>{o.cliente}</b> — {o.proximaAccion} <span style={{ color: MUTED }}>(era para {fmtDateTime(o.fechaProximaAccion)})</span>
                        </div>
                      </Card>
                    ))}
                    {!crmRecordatorios.atrasadas.length && <div style={{ fontSize: 12.5, color: MUTED }}>Sin pendientes atrasados.</div>}
                  </div>
                </div>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 8, color: WARN }}>Para hoy ({crmRecordatorios.hoy.length})</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    {crmRecordatorios.hoy.map((o) => (
                      <Card key={o.id} style={{ padding: 12, display: "flex", alignItems: "center", gap: 10, cursor: "pointer", borderColor: `${WARN}44` }} onClick={() => setCrmModal({ mode: "edit", data: o })}>
                        <div style={{ width: 32, height: 32, borderRadius: 8, background: `${WARN}16`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                          <CalendarClock size={16} color={WARN} />
                        </div>
                        <div style={{ fontSize: 13 }}>
                          <b>{o.cliente}</b> — {o.proximaAccion} <span style={{ color: MUTED }}>({fmtDateTime(o.fechaProximaAccion)})</span>
                        </div>
                      </Card>
                    ))}
                    {!crmRecordatorios.hoy.length && <div style={{ fontSize: 12.5, color: MUTED }}>Nada programado para hoy.</div>}
                  </div>
                </div>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 8 }}>Enfriándose ({crmPorReactivar.length})</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    {crmPorReactivar.map((o) => (
                      <Card key={o.id} style={{ padding: 12, display: "flex", alignItems: "center", gap: 10, cursor: "pointer" }} onClick={() => setCrmModal({ mode: "edit", data: o })}>
                        <div style={{ width: 32, height: 32, borderRadius: 8, background: `${MUTED}16`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                          <Clock size={16} color={MUTED} />
                        </div>
                        <div style={{ fontSize: 13 }}>
                          <b>{o.cliente}</b> — último contacto {fmtDate(o.fechaUltimoContacto)}
                        </div>
                      </Card>
                    ))}
                    {!crmPorReactivar.length && <div style={{ fontSize: 12.5, color: MUTED }}>Ningún cliente enfriándose por ahora.</div>}
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {tab === "ventas" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
              <div style={{ position: "relative", width: 280 }}>
                <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: MUTED }} />
                <TextInput placeholder="Buscar cliente, pedido, modelo…" value={search} onChange={(e) => setSearch(e.target.value)} style={{ paddingLeft: 32 }} />
              </div>
              <Button icon={Plus} onClick={() => setModal({ mode: "new" })}>Nuevo pedido</Button>
            </div>

            <Card style={{ overflow: "hidden" }}>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: PAPER, textAlign: "left" }}>
                      {["Fecha de pago", "Pedido", "Cliente", "Modelo", "Tipo", "Monto pagado", "Saldo del pedido", "Destino", ""].map((h) => (
                        <th key={h} style={{ padding: "10px 12px", fontSize: 11, color: MUTED, fontWeight: 700, textTransform: "uppercase", whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {movimientosOrdenados.map((m) => {
                      const venta = ventas.find((v) => v.id === m.ventaId);
                      if (!venta) return null;
                      const saldo = getSaldoPendiente(venta);
                      const liquidado = getLiquidado(venta);
                      return (
                        <tr key={m.id} style={{ borderTop: `1px solid ${LINE}` }}>
                          <td style={{ padding: "10px 12px", whiteSpace: "nowrap", color: MUTED }}>{fmtDate(m.fecha)}</td>
                          <td style={{ padding: "10px 12px", fontWeight: 600 }}>{m.numeroPedido}</td>
                          <td style={{ padding: "10px 12px" }}>
                            <div style={{ fontWeight: 600 }}>{m.cliente}</div>
                            <div style={{ fontSize: 11.5, color: MUTED }}>{m.empresa}</div>
                          </td>
                          <td style={{ padding: "10px 12px" }}>{m.modelo}</td>
                          <td style={{ padding: "10px 12px", whiteSpace: "nowrap" }}>
                            <Pill color={m.tipo === "Finiquito" ? GOOD : m.tipo === "Anticipo" ? ACCENT : WARN}>{m.tipo}</Pill>
                          </td>
                          <td style={{ padding: "10px 12px", whiteSpace: "nowrap", fontWeight: 700 }}>{fmtMoney(m.monto)}</td>
                          <td style={{ padding: "10px 12px", whiteSpace: "nowrap" }}>
                            {liquidado
                              ? <Pill color={GOOD}>Liquidado</Pill>
                              : <span style={{ color: saldo > 0 ? WARN : MUTED, fontWeight: 600 }}>{fmtMoney(saldo)}</span>}
                          </td>
                          <td style={{ padding: "10px 12px", whiteSpace: "nowrap" }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 4 }}><MapPin size={12} color={MUTED} />{m.ciudadDestino}, {m.estadoDestino}</div>
                          </td>
                          <td style={{ padding: "10px 12px" }}>
                            <div style={{ display: "flex", gap: 6 }}>
                              <button onClick={() => setQuoteModal(venta)} title="Generar cotización" style={{ border: "none", background: "transparent", cursor: "pointer", color: MUTED }}><FileText size={15} /></button>
                              <button onClick={() => setModal({ mode: "edit", data: venta })} title="Ver / editar pedido y registrar otro pago" style={{ border: "none", background: "transparent", cursor: "pointer", color: MUTED }}><Pencil size={15} /></button>
                              <button onClick={() => setConfirmDelete(venta.id)} title="Eliminar el pedido completo" style={{ border: "none", background: "transparent", cursor: "pointer", color: BAD }}><Trash2 size={15} /></button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                {!movimientosOrdenados.length && <EmptyState icon={ClipboardList} title="Sin ventas" sub="Cuando un pedido tenga su primer pago registrado, aparecerá aquí automáticamente." />}
              </div>
            </Card>
          </div>
        )}

        {tab === "clientes" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <div style={{ position: "relative", width: 280 }}>
              <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: MUTED }} />
              <TextInput placeholder="Buscar cliente, empresa, teléfono…" value={searchClientes} onChange={(e) => setSearchClientes(e.target.value)} style={{ paddingLeft: 32 }} />
            </div>
            {searchClientes.trim() && (
              <div style={{ fontSize: 12, color: MUTED }}>
                {clientesAgg.filter((c) => [c.cliente, c.empresa, c.telefono].some((f) => (f || "").toLowerCase().includes(searchClientes.toLowerCase()))).length
                  ? null
                  : `Ningún cliente registrado con "${searchClientes}".`}
              </div>
            )}
          <Card style={{ overflow: "hidden" }}>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ background: PAPER, textAlign: "left" }}>
                    {["Cliente", "Empresa", "Pedidos", "Total comprado (sin IVA)", "Ticket promedio", "Última compra", "Contacto"].map((h) => (
                      <th key={h} style={{ padding: "10px 12px", fontSize: 11, color: MUTED, fontWeight: 700, textTransform: "uppercase", whiteSpace: "nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {clientesAgg
                    .filter((c) => !searchClientes.trim() || [c.cliente, c.empresa, c.telefono].some((f) => (f || "").toLowerCase().includes(searchClientes.toLowerCase())))
                    .map((c, i) => (
                    <tr key={c.cliente} style={{ borderTop: `1px solid ${LINE}` }}>
                      <td style={{ padding: "10px 12px", fontWeight: 600 }}>
                        {i === 0 && <Pill color={GOOD} bg={`${GOOD}18`} style={{ marginRight: 6 }}>Top</Pill>}
                        {c.cliente}
                      </td>
                      <td style={{ padding: "10px 12px", color: MUTED }}>{c.empresa}</td>
                      <td style={{ padding: "10px 12px" }}>{c.pedidos}</td>
                      <td style={{ padding: "10px 12px", fontWeight: 600 }}>{fmtMoney(c.totalCobrado)}</td>
                      <td style={{ padding: "10px 12px", color: MUTED }}>{fmtMoney(c.pedidos ? c.totalCobrado / c.pedidos : 0)}</td>
                      <td style={{ padding: "10px 12px", color: MUTED }}>{fmtDate(c.ultima)}</td>
                      <td style={{ padding: "10px 12px", color: MUTED, whiteSpace: "nowrap" }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 4 }}><Phone size={12} />{c.telefono || "—"}</div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!clientesAgg.length && <EmptyState icon={Users} title="Sin clientes registrados aún" />}
            </div>
          </Card>
          </div>
        )}

        {tab === "modelos" && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 12 }}>
            {modelosAgg.map((m, i) => (
              <Card key={m.modelo} style={{ padding: 16 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <div style={{ fontWeight: 700 }}>{m.modelo}</div>
                  {i === 0 && <Pill color={GOOD}>Más vendido</Pill>}
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, fontSize: 13 }}>
                  <div><div style={{ color: MUTED, fontSize: 11 }}>Piezas</div><div style={{ fontWeight: 700 }}>{fmtNum(m.piezas)}</div></div>
                  <div><div style={{ color: MUTED, fontSize: 11 }}>m² vendidos</div><div style={{ fontWeight: 700 }}>{fmtNum(m.m2)}</div></div>
                  <div><div style={{ color: MUTED, fontSize: 11 }}>Ventas (sin IVA)</div><div style={{ fontWeight: 700 }}>{fmtMoney(m.ventas)}</div></div>
                  <div><div style={{ color: MUTED, fontSize: 11 }}>Clientes</div><div style={{ fontWeight: 700 }}>{m.clientes}</div></div>
                </div>
              </Card>
            ))}
            {!modelosAgg.length && <EmptyState icon={Boxes} title="Sin datos de modelos aún" />}
          </div>
        )}

        {tab === "reportes" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Card style={{ padding: 16 }} className="no-print">
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 10 }}>
                <Field label="Mes">
                  <Select value={reportFilters.mes} onChange={(e) => setReportFilters((f) => ({ ...f, mes: e.target.value }))}>
                    <option value="">Todos</option>
                    {MESES.map((m, i) => <option key={m} value={String(i + 1).padStart(2, "0")}>{m}</option>)}
                  </Select>
                </Field>
                <Field label="Año">
                  <TextInput placeholder="2026" value={reportFilters.anio} onChange={(e) => setReportFilters((f) => ({ ...f, anio: e.target.value }))} />
                </Field>
                <Field label="Cliente">
                  <TextInput placeholder="Nombre" value={reportFilters.cliente} onChange={(e) => setReportFilters((f) => ({ ...f, cliente: e.target.value }))} />
                </Field>
                <Field label="Modelo">
                  <Select value={reportFilters.modelo} onChange={(e) => setReportFilters((f) => ({ ...f, modelo: e.target.value }))}>
                    <option value="">Todos</option>{catalogos.modelos.map((m) => <option key={m}>{m}</option>)}
                  </Select>
                </Field>
                <Field label="Estado">
                  <Select value={reportFilters.estado} onChange={(e) => setReportFilters((f) => ({ ...f, estado: e.target.value }))}>
                    <option value="">Todos</option>{ESTADOS_MX.map((e) => <option key={e}>{e}</option>)}
                  </Select>
                </Field>
                <Field label="Ciudad">
                  <TextInput placeholder="Ciudad" value={reportFilters.ciudad} onChange={(e) => setReportFilters((f) => ({ ...f, ciudad: e.target.value }))} />
                </Field>
                <Field label="Tipo de envío">
                  <Select value={reportFilters.tipoEnvio} onChange={(e) => setReportFilters((f) => ({ ...f, tipoEnvio: e.target.value }))}>
                    <option value="">Todos</option><option>Local</option><option>Nacional</option>
                  </Select>
                </Field>
                <Field label="Etapa">
                  <Select value={reportFilters.etapa} onChange={(e) => setReportFilters((f) => ({ ...f, etapa: e.target.value }))}>
                    <option value="">Todas</option>{ETAPAS.map((e) => <option key={e}>{e}</option>)}
                  </Select>
                </Field>
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: 12, flexWrap: "wrap" }}>
                <Button variant="secondary" icon={Download} size="sm" onClick={exportCSV}>Exportar CSV</Button>
                <Button variant="secondary" icon={FileSpreadsheet} size="sm" onClick={exportExcel}>Exportar Excel</Button>
                <Button variant="secondary" icon={Printer} size="sm" onClick={exportPDF}>Exportar PDF</Button>
                <span style={{ fontSize: 12, color: MUTED, alignSelf: "center", marginLeft: "auto" }}>{reportRows.length} resultado(s)</span>
              </div>
            </Card>

            <Card style={{ overflow: "hidden" }}>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                  <thead>
                    <tr style={{ background: PAPER, textAlign: "left" }}>
                      {["Fecha", "Pedido", "Cliente", "Modelo", "Estado", "Ciudad", "Envío", "Etapa", "Subtotal"].map((h) => (
                        <th key={h} style={{ padding: "9px 10px", fontSize: 10.5, color: MUTED, fontWeight: 700, textTransform: "uppercase", whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {reportRows.map((v) => (
                      <tr key={v.id} style={{ borderTop: `1px solid ${LINE}` }}>
                        <td style={{ padding: "8px 10px", color: MUTED, whiteSpace: "nowrap" }}>{fmtDate(v.fecha)}</td>
                        <td style={{ padding: "8px 10px", fontWeight: 600 }}>{v.numeroPedido}</td>
                        <td style={{ padding: "8px 10px" }}>{v.cliente}</td>
                        <td style={{ padding: "8px 10px" }}>{v.modelo}</td>
                        <td style={{ padding: "8px 10px" }}>{v.estadoDestino}</td>
                        <td style={{ padding: "8px 10px" }}>{v.ciudadDestino}</td>
                        <td style={{ padding: "8px 10px" }}>{v.tipoEnvio}</td>
                        <td style={{ padding: "8px 10px" }}><Pill color={ETAPA_COLOR[v.etapa]}>{v.etapa}</Pill></td>
                        <td style={{ padding: "8px 10px", fontWeight: 600, whiteSpace: "nowrap" }}>{fmtMoney(getSubtotalPedido(v))}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {!reportRows.length && <EmptyState icon={FileBarChart} title="Sin resultados" sub="Ajusta los filtros para ver pedidos." />}
              </div>
            </Card>
          </div>
        )}

        {tab === "alertas" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {alertas.map((a, i) => (
              <Card key={i} style={{ padding: 14, display: "flex", alignItems: "center", gap: 12 }}>
                <div style={{ width: 34, height: 34, borderRadius: 9, background: `${a.color}18`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <a.icon size={17} color={a.color} />
                </div>
                <div style={{ fontSize: 13.5 }}>{a.texto}</div>
              </Card>
            ))}
            {!alertas.length && <EmptyState icon={BellRing} title="Todo en orden" sub="No hay alertas pendientes por ahora." />}
          </div>
        )}

        {tab === "catalogo" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Card style={{ padding: "12px 16px", display: "flex", alignItems: "center", gap: 10, background: `${ACCENT}0D`, border: `1px solid ${ACCENT}33` }}>
              <Layers size={16} color={ACCENT} />
              <div style={{ fontSize: 12.5, color: INK }}>
                Aquí administras las opciones que aparecen en el formulario de pedidos (modelo, color/paleta, acabado, fletera y forma de pago). Los pedidos ya guardados no cambian si editas esta lista.
              </div>
            </Card>

            <Card style={{ padding: 16 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
                <FileText size={16} color={NAVY} />
                <div style={{ fontWeight: 700, fontSize: 14 }}>Datos de tu empresa (aparecen en las cotizaciones)</div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 10 }}>
                <Field label="Nombre de la empresa">
                  <TextInput placeholder="Prolífica Soluciones" value={catalogos.empresaInfo?.nombre || ""} onChange={(e) => updateEmpresaInfo({ nombre: e.target.value })} />
                </Field>
                <Field label="Teléfono">
                  <TextInput placeholder="999 000 0000" value={catalogos.empresaInfo?.telefono || ""} onChange={(e) => updateEmpresaInfo({ telefono: e.target.value })} />
                </Field>
                <Field label="Correo">
                  <TextInput placeholder="contacto@empresa.com" value={catalogos.empresaInfo?.correo || ""} onChange={(e) => updateEmpresaInfo({ correo: e.target.value })} />
                </Field>
              </div>
            </Card>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 12 }}>
              <CatalogSection
                title="Modelos" icon={Boxes} items={catalogos.modelos}
                usageCount={(item) => ventas.filter((v) => v.modelo === item).length}
                onAdd={(v) => addCatalogItem("modelos", v)}
                onRemove={(i) => removeCatalogItem("modelos", i)}
              />
              <CatalogSection
                title="Colores / Paleta" icon={Palette} items={catalogos.colores}
                usageCount={(item) => ventas.filter((v) => v.color === item).length}
                onAdd={(v) => addCatalogItem("colores", v)}
                onRemove={(i) => removeCatalogItem("colores", i)}
              />
              <CatalogSection
                title="Acabados" icon={Wrench} items={catalogos.acabados}
                usageCount={(item) => ventas.filter((v) => v.acabado === item).length}
                onAdd={(v) => addCatalogItem("acabados", v)}
                onRemove={(i) => removeCatalogItem("acabados", i)}
              />
              <CatalogSection
                title="Fleteras" icon={Truck} items={catalogos.fleteras}
                usageCount={(item) => ventas.filter((v) => v.fletera === item).length}
                onAdd={(v) => addCatalogItem("fleteras", v)}
                onRemove={(i) => removeCatalogItem("fleteras", i)}
              />
              <CatalogSection
                title="Formas de pago" icon={CreditCard} items={catalogos.formasPago}
                usageCount={(item) => ventas.filter((v) => v.formaPago === item).length}
                onAdd={(v) => addCatalogItem("formasPago", v)}
                onRemove={(i) => removeCatalogItem("formasPago", i)}
              />
            </div>
          </div>
        )}
      </div>

      {/* MODALES */}
      {modal && (
        <SaleModal
          initial={modal.mode === "edit" ? modal.data : null}
          onClose={() => setModal(null)}
          onSave={handleSave}
          onAddPago={handleAddPago}
          catalogos={catalogos}
          clientesConocidos={clientesAgg}
          ventasExistentes={ventas}
        />
      )}

      {confirmDelete && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(14,26,42,0.45)", zIndex: 70, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
          <Card style={{ padding: 20, width: "100%", maxWidth: 360 }}>
            <div style={{ fontWeight: 700, marginBottom: 6 }}>¿Eliminar este pedido?</div>
            <div style={{ fontSize: 13, color: MUTED, marginBottom: 16 }}>Esta acción no se puede deshacer.</div>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
              <Button variant="secondary" onClick={() => setConfirmDelete(null)}>Cancelar</Button>
              <Button variant="danger" icon={Trash2} onClick={() => handleDelete(confirmDelete)}>Eliminar</Button>
            </div>
          </Card>
        </div>
      )}

      {quoteModal && (
        <QuoteModal venta={quoteModal} empresaInfo={catalogos.empresaInfo} onClose={() => setQuoteModal(null)} />
      )}

      {crmModal && (
        <OpportunityModal
          initial={crmModal.mode === "edit" ? crmModal.data : null}
          onClose={() => setCrmModal(null)}
          onSave={crmSaveOpportunity}
          onDelete={(id) => { setCrmModal(null); setCrmConfirmDelete(id); }}
          acciones={accionesCRM}
          onAddAccion={crmAddAccion}
        />
      )}

      {crmConfirmDelete && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(14,26,42,0.45)", zIndex: 70, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
          <Card style={{ padding: 20, width: "100%", maxWidth: 360 }}>
            <div style={{ fontWeight: 700, marginBottom: 6 }}>¿Eliminar esta oportunidad?</div>
            <div style={{ fontSize: 13, color: MUTED, marginBottom: 16 }}>Esta acción no se puede deshacer.</div>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
              <Button variant="secondary" onClick={() => setCrmConfirmDelete(null)}>Cancelar</Button>
              <Button variant="danger" icon={Trash2} onClick={() => crmDeleteOpportunity(crmConfirmDelete)}>Eliminar</Button>
            </div>
          </Card>
        </div>
      )}
    </div>
  );
}
