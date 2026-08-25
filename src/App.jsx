import React, { useEffect, useState } from "react";
import { auth, db } from "./firebase";
import {
  createUserWithEmailAndPassword,
  signInWithEmailAndPassword,
  signOut,
  onAuthStateChanged,
} from "firebase/auth";
import {
  collection,
  doc,
  setDoc,
  addDoc,
  updateDoc,
  deleteDoc,
  onSnapshot,
  getDoc,
} from "firebase/firestore";

const CATEGORIAS = ["Sub-6", "Sub-8", "Sub-10", "Sub-12", "Sub-14", "Sub-16", "Sub-16 Fem", "Honor", "Honor Fem"];

// Roles que existen: "admin", "tesorera", "profesor", "sin-rol"
// Qué pestañas ve cada rol:
const TABS_BY_ROLE = {
  admin: ["fichas", "convocatorias", "asistencia", "partidos", "series", "profesores", "caja"],
  tesorera: ["fichas", "convocatorias", "asistencia", "partidos", "series", "profesores", "caja"],
  profesor: ["fichas", "convocatorias", "asistencia", "partidos", "series"],
  "sin-rol": [],
};

export default function App() {
  const [user, setUser] = useState(null);
  const [role, setRole] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (u) => {
      setUser(u);
      if (u) {
        const snap = await getDoc(doc(db, "usuarios", u.uid));
        setRole(snap.exists() ? snap.data().role : "sin-rol");
      } else {
        setRole(null);
      }
      setLoading(false);
    });
    return () => unsub();
  }, []);

  if (loading) return <div className="main">Cargando...</div>;
  if (!user) return <Login />;
  if (role === "sin-rol")
    return (
      <div className="main">
        <div className="card">
          Tu cuenta ({user.email}) todavía no tiene un rol asignado. Pídele a la
          administradora que te asigne rol (admin / tesorera / profesor) en el
          panel de Firestore, colección "usuarios", documento con tu UID: <b>{user.uid}</b>
        </div>
        <button className="secondary" onClick={() => signOut(auth)}>Cerrar sesión</button>
      </div>
    );

  return <Dashboard user={user} role={role} />;
}

function Login() {
  const [mode, setMode] = useState("login"); // login | signup
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  async function submit(e) {
    e.preventDefault();
    setError("");
    try {
      if (mode === "login") {
        await signInWithEmailAndPassword(auth, email, password);
      } else {
        const cred = await createUserWithEmailAndPassword(auth, email, password);
        // Crea el documento de usuario con rol "sin-rol" por defecto.
        // La administradora debe entrar a Firestore y cambiarle el campo "role"
        // a "admin", "tesorera" o "profesor" según corresponda.
        await setDoc(doc(db, "usuarios", cred.user.uid), {
          email,
          role: "sin-rol",
        });
      }
    } catch (err) {
      setError(err.message);
    }
  }

  return (
    <div className="login-box">
      <h2>Academia Curanipe</h2>
      <form onSubmit={submit}>
        <input placeholder="Correo" type="email" value={email} onChange={(e) => setEmail(e.target.value)} required />
        <input placeholder="Contraseña" type="password" value={password} onChange={(e) => setPassword(e.target.value)} required />
        {error && <div className="error">{error}</div>}
        <button className="primary" type="submit">{mode === "login" ? "Entrar" : "Crear cuenta"}</button>
      </form>
      <p className="muted" style={{ marginTop: 12 }}>
        {mode === "login" ? (
          <>¿No tienes cuenta? <a href="#" onClick={(e) => { e.preventDefault(); setMode("signup"); }}>Crear una</a></>
        ) : (
          <>¿Ya tienes cuenta? <a href="#" onClick={(e) => { e.preventDefault(); setMode("login"); }}>Entrar</a></>
        )}
      </p>
    </div>
  );
}

function Dashboard({ user, role }) {
  const tabsDisponibles = TABS_BY_ROLE[role] || [];
  const [tab, setTab] = useState(tabsDisponibles[0] || "fichas");
  const labels = { fichas: "Fichas", convocatorias: "Convocatorias", asistencia: "Asistencia", partidos: "Partidos", series: "Resumen por serie", profesores: "Profesores", caja: "Caja general" };

  return (
    <div>
      <div className="header">
        <h1>Academia Curanipe — Gestión</h1>
        <div className="nav">
          {tabsDisponibles.map((t) => (
            <button key={t} className={tab === t ? "active" : ""} onClick={() => setTab(t)}>{labels[t]}</button>
          ))}
          <button onClick={() => signOut(auth)}>Salir ({role})</button>
        </div>
      </div>
      <div className="main">
        {tab === "fichas" && <Fichas />}
        {tab === "convocatorias" && <Convocatorias />}
        {tab === "asistencia" && <Asistencia />}
        {tab === "partidos" && <Partidos />}
        {tab === "series" && <Series />}
        {tab === "profesores" && <Profesores />}
        {tab === "caja" && <Caja />}
      </div>
    </div>
  );
}

function emptyForm() {
  return {
    nombre: "", apellido: "", rut: "", categoria: CATEGORIAS[0], talla: "", apoderadoNombre: "", apoderadoTelefono: "",
    fechaNacimiento: "", matriculaPagada: false, matriculaMonto: "",
  };
}

function normalizaCategoria(txt) {
  if (!txt) return CATEGORIAS[0];
  const t = String(txt).toLowerCase().replace(/\s+/g, "");
  const match = CATEGORIAS.find((c) => t.includes(c.toLowerCase().replace("-", "")));
  return match || CATEGORIAS[0];
}

function buscaCol(row, palabrasClave) {
  const keys = Object.keys(row);
  for (const k of keys) {
    const norm = k.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
    if (palabrasClave.some((p) => norm === p || norm.includes(p))) return row[k];
  }
  return "";
}

function Fichas() {
  const [players, setPlayers] = useState([]);
  const [form, setForm] = useState(emptyForm());
  const [editingId, setEditingId] = useState(null);
  const [msg, setMsg] = useState("");
  const [importando, setImportando] = useState(false);

  async function handleImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    setImportando(true);
    try {
      const XLSX = await import("xlsx");
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data);
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      let agregadas = 0;
      let saltadas = 0;
      for (const row of rows) {
        const nombre = buscaCol(row, ["nombre", "nombres"]);
        if (!String(nombre).trim()) { saltadas++; continue; }
        const apellido = buscaCol(row, ["apellido", "apellidos"]);
        const rut = buscaCol(row, ["rut"]);
        const fechaRaw = buscaCol(row, ["fecha de nacimiento", "fechanacimiento", "fecha nac", "nacimiento"]);
        let fechaNacimiento = "";
        if (fechaRaw instanceof Date) {
          fechaNacimiento = fechaRaw.toISOString().slice(0, 10);
        } else if (typeof fechaRaw === "string" && fechaRaw.includes("/")) {
          const [d, m, y] = fechaRaw.split("/");
          if (d && m && y) fechaNacimiento = `${y.length === 2 ? "20" + y : y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
        } else if (fechaRaw) {
          fechaNacimiento = String(fechaRaw);
        }
        const categoria = normalizaCategoria(buscaCol(row, ["categoria"]));
        await addDoc(collection(db, "jugadoras"), {
          ...emptyForm(),
          nombre: String(nombre).trim(),
          apellido: String(apellido).trim(),
          rut: String(rut).trim(),
          fechaNacimiento,
          categoria,
        });
        agregadas++;
      }
      setMsg(`Se importaron ${agregadas} jugadoras${saltadas ? ` (se saltaron ${saltadas} filas sin nombre)` : ""}. Revisa la lista y completa apoderado/teléfono si faltan.`);
    } catch (err) {
      setMsg("Error al importar: " + err.message);
    }
    setImportando(false);
    e.target.value = "";
  }

  useEffect(() => {
    const unsub = onSnapshot(collection(db, "jugadoras"), (snap) => {
      setPlayers(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });
    return () => unsub();
  }, []);

  async function submit() {
    if (!form.nombre.trim()) { setMsg("Falta el nombre."); return; }
    setMsg("");
    if (editingId) {
      await updateDoc(doc(db, "jugadoras", editingId), form);
      setEditingId(null);
    } else {
      await addDoc(collection(db, "jugadoras"), form);
    }
    setForm(emptyForm());
  }
  function edit(p) { setForm({ ...emptyForm(), ...p }); setEditingId(p.id); }
  async function remove(id) { await deleteDoc(doc(db, "jugadoras", id)); }

  return (
    <div>
      <div className="card">
        <div className="muted" style={{ marginBottom: 8 }}>Importar jugadoras desde tu Excel (columnas: Nombre, Apellido, Rut, Fecha de nacimiento, Categoria)</div>
        <input type="file" accept=".xlsx,.xls,.csv" onChange={handleImport} disabled={importando} />
        {importando && <div className="muted">Importando...</div>}
      </div>

      <div className="card">
        {msg && <div className="error">{msg}</div>}
        <div className="row">
          <input placeholder="Nombre" value={form.nombre} onChange={(e) => setForm({ ...form, nombre: e.target.value })} />
          <input placeholder="Apellido" value={form.apellido} onChange={(e) => setForm({ ...form, apellido: e.target.value })} />
          <input placeholder="RUT" value={form.rut} onChange={(e) => setForm({ ...form, rut: e.target.value })} />
          <select value={form.categoria} onChange={(e) => setForm({ ...form, categoria: e.target.value })}>
            {CATEGORIAS.map((c) => <option key={c}>{c}</option>)}
          </select>
          <input type="date" placeholder="Fecha nacimiento" value={form.fechaNacimiento} onChange={(e) => setForm({ ...form, fechaNacimiento: e.target.value })} />
          <input placeholder="Talla" value={form.talla} onChange={(e) => setForm({ ...form, talla: e.target.value })} />
          <input placeholder="Apoderado" value={form.apoderadoNombre} onChange={(e) => setForm({ ...form, apoderadoNombre: e.target.value })} />
          <input placeholder="Teléfono (56912345678)" value={form.apoderadoTelefono} onChange={(e) => setForm({ ...form, apoderadoTelefono: e.target.value })} />
        </div>
        <div className="row" style={{ marginTop: 8 }}>
          <button className="primary" onClick={submit}>{editingId ? "Guardar cambios" : "Agregar jugadora"}</button>
          {editingId && <button className="secondary" onClick={() => { setEditingId(null); setForm(emptyForm()); }}>Cancelar</button>}
        </div>
      </div>

      <div className="card">
        {players.length === 0 && <div className="muted">No hay jugadoras registradas.</div>}
        {players.map((p) => (
          <div key={p.id} className="list-item">
            <div>
              <b>{p.nombre} {p.apellido}</b> <span className="muted">— {p.categoria}</span>
              <div className="muted">RUT: {p.rut || "-"} · Nac: {p.fechaNacimiento || "-"}</div>
              <div className="muted">Apoderado: {p.apoderadoNombre} {p.apoderadoTelefono && `(${p.apoderadoTelefono})`}</div>
            </div>
            <div>
              <button className="secondary" onClick={() => edit(p)}>Editar</button>{" "}
              <button className="secondary" onClick={() => remove(p.id)}>Eliminar</button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function waLink(phone, msg) {
  const p = (phone || "").replace(/[^0-9]/g, "");
  if (!p) return null;
  return `https://wa.me/${p}?text=${encodeURIComponent(msg)}`;
}

function Convocatorias() {
  const [players, setPlayers] = useState([]);
  const [convs, setConvs] = useState([]);
  const [categoria, setCategoria] = useState(CATEGORIAS[0]);
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [lugar, setLugar] = useState("");
  const [hora, setHora] = useState("");
  const [tipo, setTipo] = useState("Entrenamiento");
  const [convId, setConvId] = useState(null);
  const [respuestas, setRespuestas] = useState({});

  useEffect(() => {
    const unsub1 = onSnapshot(collection(db, "jugadoras"), (snap) => {
      setPlayers(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });
    const unsub2 = onSnapshot(collection(db, "convocatorias"), (snap) => {
      setConvs(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });
    return () => { unsub1(); unsub2(); };
  }, []);

  const catPlayers = players.filter((p) => p.categoria === categoria);
  const key = `${fecha}__${categoria}`;
  const convExistente = convs.find((c) => c.fecha === fecha && c.categoria === categoria);

  useEffect(() => {
    if (convExistente) {
      setConvId(convExistente.id);
      setLugar(convExistente.lugar || "");
      setHora(convExistente.hora || "");
      setTipo(convExistente.tipo || "Entrenamiento");
      setRespuestas(convExistente.respuestas || {});
    } else {
      setConvId(null);
      setRespuestas({});
    }
    // eslint-disable-next-line
  }, [key]);

  async function crearOActualizar() {
    const datos = { categoria, fecha, lugar, hora, tipo, respuestas };
    if (convId) {
      await updateDoc(doc(db, "convocatorias", convId), datos);
    } else {
      const ref = await addDoc(collection(db, "convocatorias"), { ...datos, creadoEn: new Date().toISOString() });
      setConvId(ref.id);
    }
  }

  async function setEstado(playerId, estado) {
    const nuevas = { ...respuestas, [playerId]: estado };
    setRespuestas(nuevas);
    if (convId) {
      await updateDoc(doc(db, "convocatorias", convId), { respuestas: nuevas });
    }
  }

  function mensajeConvocatoria(p) {
    return `Hola${p.apoderadoNombre ? " " + p.apoderadoNombre : ""}, te convocamos a ${p.nombre} para el ${fecha}${hora ? " a las " + hora : ""} (${categoria}, ${tipo})${lugar ? " en " + lugar : ""}. ¿Podrá asistir? — Academia Curanipe`;
  }

  const confirmadas = catPlayers.filter((p) => respuestas[p.id] === "Confirmado").length;

  async function eliminarConv(id) {
    await deleteDoc(doc(db, "convocatorias", id));
    if (id === convId) { setConvId(null); setRespuestas({}); }
  }

  function cargarConv(c) {
    setCategoria(c.categoria);
    setFecha(c.fecha);
  }

  const historial = convs
    .filter((c) => c.categoria === categoria)
    .sort((a, b) => (a.fecha < b.fecha ? 1 : -1))
    .slice(0, 15);

  return (
    <div>
      <div className="card">
        <div className="row">
          <select value={categoria} onChange={(e) => setCategoria(e.target.value)}>
            {CATEGORIAS.map((c) => <option key={c}>{c}</option>)}
          </select>
          <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} />
          <input placeholder="Lugar" value={lugar} onChange={(e) => setLugar(e.target.value)} />
          <input type="time" value={hora} onChange={(e) => setHora(e.target.value)} />
          <select value={tipo} onChange={(e) => setTipo(e.target.value)}>
            <option>Entrenamiento</option>
            <option>Partido</option>
            <option>Actividad extra</option>
          </select>
        </div>
        <button className="primary" style={{ marginTop: 10 }} onClick={crearOActualizar}>
          {convId ? "Guardar datos de la convocatoria" : "Crear convocatoria"}
        </button>
        <div className="muted" style={{ marginTop: 6 }}>
          Envía el mensaje por WhatsApp y anota manualmente lo que responda cada apoderado (no se captura solo).
        </div>
      </div>

      <div className="card">
        <div className="muted" style={{ marginBottom: 8 }}>{confirmadas}/{catPlayers.length} confirmadas</div>
        {catPlayers.length === 0 && <div className="muted">No hay jugadoras en esta categoría.</div>}
        {catPlayers.map((p) => {
          const link = waLink(p.apoderadoTelefono, mensajeConvocatoria(p));
          return (
            <div key={p.id} className="list-item">
              <div>{p.nombre} {p.apellido}</div>
              <div className="row">
                <select value={respuestas[p.id] || "Pendiente"} onChange={(e) => setEstado(p.id, e.target.value)}>
                  <option>Pendiente</option>
                  <option>Confirmado</option>
                  <option>No asiste</option>
                </select>
                {link ? (
                  <a href={link} target="_blank" rel="noopener noreferrer" className="secondary" style={{ textDecoration: "none", display: "inline-block" }}>Enviar WhatsApp</a>
                ) : (
                  <span className="muted">Sin teléfono</span>
                )}
              </div>
            </div>
          );
        })}
      </div>

      <div className="card">
        <div className="muted" style={{ marginBottom: 8 }}>Últimas convocatorias de {categoria}</div>
        {historial.length === 0 && <div className="muted">Sin convocatorias registradas.</div>}
        {historial.map((c) => {
          const conf = Object.values(c.respuestas || {}).filter((v) => v === "Confirmado").length;
          const total = players.filter((p) => p.categoria === c.categoria).length;
          return (
            <div key={c.id} className="list-item">
              <div>
                <b>{c.fecha}</b> · {c.tipo} {c.lugar && `· ${c.lugar}`} {c.hora && `· ${c.hora}`}
                <div className="muted">{conf}/{total} confirmadas</div>
              </div>
              <div>
                <button className="secondary" onClick={() => cargarConv(c)}>Abrir</button>{" "}
                <button className="secondary" onClick={() => eliminarConv(c.id)}>Eliminar</button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function Asistencia() {
  const [players, setPlayers] = useState([]);
  const [sesiones, setSesiones] = useState([]);
  const [convs, setConvs] = useState([]);
  const [categoria, setCategoria] = useState(CATEGORIAS[0]);
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [lugar, setLugar] = useState("");
  const [hora, setHora] = useState("");
  const [tipo, setTipo] = useState("Entrenamiento");
  const [profesor, setProfesor] = useState("");
  const [presentes, setPresentes] = useState({});
  const [msg, setMsg] = useState("");
  const [editingId, setEditingId] = useState(null);

  useEffect(() => {
    const unsub1 = onSnapshot(collection(db, "jugadoras"), (snap) => {
      setPlayers(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });
    const unsub2 = onSnapshot(collection(db, "asistencia"), (snap) => {
      setSesiones(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });
    const unsub3 = onSnapshot(collection(db, "convocatorias"), (snap) => {
      setConvs(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });
    return () => { unsub1(); unsub2(); unsub3(); };
  }, []);

  // Si no estamos editando una sesión ya guardada, y existe una convocatoria
  // para esa misma fecha/categoría, precarga lugar/hora/tipo desde ahí.
  useEffect(() => {
    if (editingId) return;
    const conv = convs.find((c) => c.categoria === categoria && c.fecha === fecha);
    if (conv) {
      setLugar(conv.lugar || "");
      setHora(conv.hora || "");
      setTipo(conv.tipo || "Entrenamiento");
    }
    // eslint-disable-next-line
  }, [categoria, fecha, convs]);

  const catPlayers = players.filter((p) => p.categoria === categoria);

  function toggle(id) {
    setPresentes((prev) => ({ ...prev, [id]: !prev[id] }));
  }

  async function guardar() {
    if (!profesor.trim()) { setMsg("Falta el nombre del profesor a cargo."); return; }
    setMsg("");
    const listaPresentes = catPlayers.filter((p) => presentes[p.id]).map((p) => p.id);
    const datos = {
      categoria, fecha, lugar, hora, tipo, profesor,
      presentes: listaPresentes,
      totalJugadoras: catPlayers.length,
    };
    if (editingId) {
      await updateDoc(doc(db, "asistencia", editingId), datos);
      setEditingId(null);
      setMsg("Sesión actualizada correctamente.");
    } else {
      await addDoc(collection(db, "asistencia"), { ...datos, creadoEn: new Date().toISOString() });
      setMsg("Sesión guardada correctamente.");
    }
    setPresentes({});
  }

  function editarSesion(s) {
    setEditingId(s.id);
    setCategoria(s.categoria);
    setFecha(s.fecha);
    setLugar(s.lugar || "");
    setHora(s.hora || "");
    setTipo(s.tipo);
    setProfesor(s.profesor || "");
    const marcadas = {};
    (s.presentes || []).forEach((id) => (marcadas[id] = true));
    setPresentes(marcadas);
    setMsg("Editando sesión del " + s.fecha + " — corrige lo que necesites y aprieta Guardar.");
  }

  function cancelarEdicion() {
    setEditingId(null);
    setPresentes({});
    setMsg("");
  }

  async function eliminarSesion(id) {
    await deleteDoc(doc(db, "asistencia", id));
  }

  const sesionesCategoria = sesiones
    .filter((s) => s.categoria === categoria)
    .sort((a, b) => (a.fecha < b.fecha ? 1 : -1))
    .slice(0, 15);

  return (
    <div>
      <div className="card">
        {msg && <div className={msg.includes("Falta") ? "error" : "muted"}>{msg}</div>}
        <div className="row">
          <select value={categoria} onChange={(e) => { setCategoria(e.target.value); setPresentes({}); }}>
            {CATEGORIAS.map((c) => <option key={c}>{c}</option>)}
          </select>
          <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} />
          <input placeholder="Lugar" value={lugar} onChange={(e) => setLugar(e.target.value)} />
          <input type="time" value={hora} onChange={(e) => setHora(e.target.value)} />
          <select value={tipo} onChange={(e) => setTipo(e.target.value)}>
            <option>Entrenamiento</option>
            <option>Partido</option>
            <option>Actividad extra</option>
            <option>Suspendido</option>
          </select>
          <input placeholder="Profesor a cargo" value={profesor} onChange={(e) => setProfesor(e.target.value)} />
        </div>
      </div>

      <div className="card">
        <div className="muted" style={{ marginBottom: 8 }}>
          Marca presentes ({catPlayers.filter((p) => presentes[p.id]).length}/{catPlayers.length})
        </div>
        {catPlayers.length === 0 && <div className="muted">No hay jugadoras en esta categoría.</div>}
        {catPlayers.map((p) => (
          <label key={p.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 0" }}>
            <input type="checkbox" checked={!!presentes[p.id]} onChange={() => toggle(p.id)} />
            {p.nombre} {p.apellido}
          </label>
        ))}
        <button className="primary" style={{ marginTop: 10 }} onClick={guardar}>{editingId ? "Guardar cambios" : "Guardar asistencia"}</button>
        {editingId && <button className="secondary" style={{ marginTop: 10, marginLeft: 8 }} onClick={cancelarEdicion}>Cancelar edición</button>}
      </div>

      <div className="card">
        <div className="muted" style={{ marginBottom: 8 }}>Últimas sesiones de {categoria}</div>
        {sesionesCategoria.length === 0 && <div className="muted">Sin sesiones registradas.</div>}
        {sesionesCategoria.map((s) => (
          <div key={s.id} className="list-item">
            <div>
              <b>{s.fecha}</b> · {s.tipo} {s.lugar && `· ${s.lugar}`} {s.hora && `· ${s.hora}`}
              <div className="muted">Profesor: {s.profesor} · {s.presentes?.length || 0}/{s.totalJugadoras} presentes</div>
            </div>
            <div>
              <button className="secondary" onClick={() => editarSesion(s)}>Editar</button>{" "}
              <button className="secondary" onClick={() => eliminarSesion(s.id)}>Eliminar</button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function emptyProfesorForm() {
  return { nombre: "", telefono: "", categorias: [] };
}

function Partidos() {
  const [partidos, setPartidos] = useState([]);
  const [players, setPlayers] = useState([]);
  const [asistencias, setAsistencias] = useState([]);
  const [categoria, setCategoria] = useState(CATEGORIAS[0]);
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [rival, setRival] = useState("");
  const [tipo, setTipo] = useState("Liga");
  const [golesPropios, setGolesPropios] = useState("");
  const [golesRival, setGolesRival] = useState("");
  const [profesor, setProfesor] = useState("");
  const [evaluacion, setEvaluacion] = useState("");
  const [jugadoras, setJugadoras] = useState({}); // { playerId: { jugo, goles, asistencias } }
  const [editingId, setEditingId] = useState(null);
  const [msg, setMsg] = useState("");

  useEffect(() => {
    const unsub = onSnapshot(collection(db, "partidos"), (snap) => {
      setPartidos(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });
    const unsub2 = onSnapshot(collection(db, "jugadoras"), (snap) => {
      setPlayers(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });
    const unsub3 = onSnapshot(collection(db, "asistencia"), (snap) => {
      setAsistencias(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });
    return () => { unsub(); unsub2(); unsub3(); };
  }, []);

  // Si no estamos editando un partido ya guardado, busca la sesión de Asistencia
  // (tipo Partido) para esa misma fecha/categoría y precarga quiénes jugaron y el profesor.
  useEffect(() => {
    if (editingId) return;
    const yaExistePartido = partidos.some((p) => p.categoria === categoria && p.fecha === fecha);
    if (yaExistePartido) return; // si ya hay un partido guardado para esa fecha, no lo pisamos
    const sesion = asistencias.find((s) => s.categoria === categoria && s.fecha === fecha && s.tipo === "Partido");
    if (sesion) {
      const nuevas = {};
      (sesion.presentes || []).forEach((id) => { nuevas[id] = { jugo: true, goles: 0, asistencias: 0 }; });
      setJugadoras(nuevas);
      if (sesion.profesor) setProfesor(sesion.profesor);
    }
    // eslint-disable-next-line
  }, [categoria, fecha, asistencias, partidos]);

  const catPlayers = players.filter((p) => p.categoria === categoria);

  function datosJugadora(id) {
    return jugadoras[id] || { jugo: false, goles: 0, asistencias: 0 };
  }
  function setJugo(id, val) {
    setJugadoras((prev) => ({ ...prev, [id]: { ...datosJugadora(id), jugo: val } }));
  }
  function setGoles(id, val) {
    setJugadoras((prev) => ({ ...prev, [id]: { ...datosJugadora(id), goles: Number(val) || 0 } }));
  }
  function setAsist(id, val) {
    setJugadoras((prev) => ({ ...prev, [id]: { ...datosJugadora(id), asistencias: Number(val) || 0 } }));
  }

  function limpiar() {
    setFecha(new Date().toISOString().slice(0, 10));
    setRival(""); setTipo("Liga"); setGolesPropios(""); setGolesRival("");
    setProfesor(""); setEvaluacion(""); setJugadoras({}); setEditingId(null);
  }

  async function guardar() {
    if (!rival.trim()) { setMsg("Falta el rival."); return; }
    setMsg("");
    const datos = {
      categoria, fecha, rival, tipo,
      golesPropios: Number(golesPropios) || 0,
      golesRival: Number(golesRival) || 0,
      profesor, evaluacion, jugadoras,
    };
    if (editingId) {
      await updateDoc(doc(db, "partidos", editingId), datos);
    } else {
      await addDoc(collection(db, "partidos"), datos);
    }
    limpiar();
  }

  function editar(p) {
    setEditingId(p.id);
    setCategoria(p.categoria); setFecha(p.fecha); setRival(p.rival); setTipo(p.tipo);
    setGolesPropios(p.golesPropios); setGolesRival(p.golesRival);
    setProfesor(p.profesor || ""); setEvaluacion(p.evaluacion || "");
    setJugadoras(p.jugadoras || {});
  }
  async function eliminar(id) { await deleteDoc(doc(db, "partidos", id)); }

  const listaCategoria = partidos.filter((p) => p.categoria === categoria).sort((a, b) => (a.fecha < b.fecha ? 1 : -1));

  return (
    <div>
      <div className="card">
        {msg && <div className="error">{msg}</div>}
        <div className="row">
          <select value={categoria} onChange={(e) => setCategoria(e.target.value)}>
            {CATEGORIAS.map((c) => <option key={c}>{c}</option>)}
          </select>
          <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} />
          <input placeholder="Rival" value={rival} onChange={(e) => setRival(e.target.value)} />
          <select value={tipo} onChange={(e) => setTipo(e.target.value)}>
            <option>Liga</option>
            <option>Amistoso</option>
          </select>
          <input type="number" placeholder="Goles propios" value={golesPropios} onChange={(e) => setGolesPropios(e.target.value)} style={{ width: 90 }} />
          <input type="number" placeholder="Goles rival" value={golesRival} onChange={(e) => setGolesRival(e.target.value)} style={{ width: 90 }} />
          <input placeholder="Profesor" value={profesor} onChange={(e) => setProfesor(e.target.value)} />
        </div>
        <textarea placeholder="Evaluación general del encuentro" value={evaluacion} onChange={(e) => setEvaluacion(e.target.value)} style={{ width: "100%", marginTop: 8 }} rows={2} />

        <div className="muted" style={{ margin: "12px 0 6px" }}>Jugadoras: marca quién jugó y sus goles/asistencias (deja sin marcar a las que no llegaron). Si ya registraste este partido en Asistencia, se precargan solas.</div>
        {catPlayers.length === 0 && <div className="muted">No hay jugadoras en esta categoría.</div>}
        {catPlayers.map((p) => {
          const d = datosJugadora(p.id);
          return (
            <div key={p.id} className="row" style={{ padding: "4px 0", borderBottom: "1px solid #f0f0f0" }}>
              <label style={{ display: "flex", alignItems: "center", gap: 6, minWidth: 160 }}>
                <input type="checkbox" checked={d.jugo} onChange={(e) => setJugo(p.id, e.target.checked)} />
                {p.nombre} {p.apellido}
              </label>
              {d.jugo && (
                <>
                  <input type="number" min="0" placeholder="Goles" value={d.goles} onChange={(e) => setGoles(p.id, e.target.value)} style={{ width: 80 }} />
                  <input type="number" min="0" placeholder="Asistencias" value={d.asistencias} onChange={(e) => setAsist(p.id, e.target.value)} style={{ width: 100 }} />
                </>
              )}
            </div>
          );
        })}

        <div className="row" style={{ marginTop: 8 }}>
          <button className="primary" onClick={guardar}>{editingId ? "Guardar cambios" : "Registrar partido"}</button>
          {editingId && <button className="secondary" onClick={limpiar}>Cancelar</button>}
        </div>
      </div>

      <div className="card">
        {listaCategoria.length === 0 && <div className="muted">No hay partidos registrados para {categoria}.</div>}
        {listaCategoria.map((p) => {
          const jugaron = Object.entries(p.jugadoras || {}).filter(([, v]) => v.jugo);
          const goleadoras = jugaron.filter(([, v]) => v.goles > 0);
          return (
            <div key={p.id} className="list-item">
              <div>
                <b>{p.fecha}</b> · vs {p.rival} ({p.tipo}) — {p.golesPropios}-{p.golesRival}
                {p.profesor && <div className="muted">Profesor: {p.profesor}</div>}
                <div className="muted">{jugaron.length} jugadoras participaron</div>
                {goleadoras.length > 0 && (
                  <div className="muted">
                    Goleadoras: {goleadoras.map(([id, v]) => {
                      const player = players.find((pl) => pl.id === id);
                      return `${player?.nombre || "-"} (${v.goles})`;
                    }).join(", ")}
                  </div>
                )}
                {p.evaluacion && <div className="muted">"{p.evaluacion}"</div>}
              </div>
              <div>
                <button className="secondary" onClick={() => editar(p)}>Editar</button>{" "}
                <button className="secondary" onClick={() => eliminar(p.id)}>Eliminar</button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function Series() {
  const [categoria, setCategoria] = useState(CATEGORIAS[0]);
  const [players, setPlayers] = useState([]);
  const [asistencia, setAsistencia] = useState([]);
  const [convs, setConvs] = useState([]);
  const [partidos, setPartidos] = useState([]);

  useEffect(() => {
    const u1 = onSnapshot(collection(db, "jugadoras"), (s) => setPlayers(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u2 = onSnapshot(collection(db, "asistencia"), (s) => setAsistencia(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u3 = onSnapshot(collection(db, "convocatorias"), (s) => setConvs(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u4 = onSnapshot(collection(db, "partidos"), (s) => setPartidos(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    return () => { u1(); u2(); u3(); u4(); };
  }, []);

  const inscritas = players.filter((p) => p.categoria === categoria);
  const sesionesCat = asistencia.filter((s) => s.categoria === categoria);
  const totalPresentes = sesionesCat.reduce((sum, s) => sum + (s.presentes?.length || 0), 0);
  const totalPosible = sesionesCat.reduce((sum, s) => sum + (s.totalJugadoras || 0), 0);
  const pctAsistencia = totalPosible ? Math.round((totalPresentes / totalPosible) * 100) : 0;

  const convsCat = convs.filter((c) => c.categoria === categoria);
  let confirmadas = 0, totalConv = 0;
  convsCat.forEach((c) => {
    Object.values(c.respuestas || {}).forEach((v) => {
      totalConv++;
      if (v === "Confirmado") confirmadas++;
    });
  });

  const partidosCat = partidos.filter((p) => p.categoria === categoria).sort((a, b) => (a.fecha < b.fecha ? 1 : -1));

  return (
    <div>
      <div className="card">
        <select value={categoria} onChange={(e) => setCategoria(e.target.value)}>
          {CATEGORIAS.map((c) => <option key={c}>{c}</option>)}
        </select>
      </div>

      <div className="row" style={{ marginBottom: 16 }}>
        <div className="card" style={{ flex: 1, minWidth: 140 }}>
          <div className="muted">Alumnas inscritas</div>
          <div style={{ fontSize: 22, fontWeight: "bold" }}>{inscritas.length}</div>
        </div>
        <div className="card" style={{ flex: 1, minWidth: 140 }}>
          <div className="muted">Asistencia promedio</div>
          <div style={{ fontSize: 22, fontWeight: "bold" }}>{pctAsistencia}%</div>
        </div>
        <div className="card" style={{ flex: 1, minWidth: 140 }}>
          <div className="muted">Convocatorias confirmadas</div>
          <div style={{ fontSize: 22, fontWeight: "bold" }}>{confirmadas}/{totalConv}</div>
        </div>
        <div className="card" style={{ flex: 1, minWidth: 140 }}>
          <div className="muted">Partidos jugados</div>
          <div style={{ fontSize: 22, fontWeight: "bold" }}>{partidosCat.length}</div>
        </div>
      </div>

      <div className="card">
        <div className="muted" style={{ marginBottom: 8 }}>Resultados de partidos</div>
        {partidosCat.length === 0 && <div className="muted">Sin partidos registrados.</div>}
        {partidosCat.map((p) => (
          <div key={p.id} className="muted" style={{ padding: "4px 0" }}>
            {p.fecha} · vs {p.rival} ({p.tipo}) — {p.golesPropios}-{p.golesRival}
          </div>
        ))}
      </div>

      <div className="card">
        <div className="muted" style={{ marginBottom: 8 }}>Alumnas inscritas</div>
        <div className="muted">{inscritas.map((p) => `${p.nombre} ${p.apellido}`).join(", ") || "Ninguna"}</div>
      </div>

      <div className="card">
        <div className="muted">El nivel táctico de equipo se va a sumar aquí cuando armemos el módulo de diagnóstico táctico.</div>
      </div>
    </div>
  );
}

function Profesores() {
  const [profes, setProfes] = useState([]);
  const [form, setForm] = useState(emptyProfesorForm());
  const [editingId, setEditingId] = useState(null);
  const [msg, setMsg] = useState("");

  useEffect(() => {
    const unsub = onSnapshot(collection(db, "profesores"), (snap) => {
      setProfes(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });
    return () => unsub();
  }, []);

  function toggleCategoria(c) {
    setForm((f) => ({
      ...f,
      categorias: f.categorias.includes(c) ? f.categorias.filter((x) => x !== c) : [...f.categorias, c],
    }));
  }

  async function submit() {
    if (!form.nombre.trim()) { setMsg("Falta el nombre del profesor/a."); return; }
    setMsg("");
    if (editingId) {
      await updateDoc(doc(db, "profesores", editingId), form);
      setEditingId(null);
    } else {
      await addDoc(collection(db, "profesores"), form);
    }
    setForm(emptyProfesorForm());
  }
  function edit(p) { setForm({ ...emptyProfesorForm(), ...p }); setEditingId(p.id); }
  async function remove(id) { await deleteDoc(doc(db, "profesores", id)); }

  return (
    <div>
      <div className="card">
        {msg && <div className="error">{msg}</div>}
        <div className="row">
          <input placeholder="Nombre completo" value={form.nombre} onChange={(e) => setForm({ ...form, nombre: e.target.value })} />
          <input placeholder="Teléfono (56912345678)" value={form.telefono} onChange={(e) => setForm({ ...form, telefono: e.target.value })} />
        </div>
        <div className="muted" style={{ margin: "10px 0 4px" }}>Categorías en las que hace clases:</div>
        <div className="row">
          {CATEGORIAS.map((c) => (
            <label key={c} style={{ display: "flex", alignItems: "center", gap: 4 }}>
              <input type="checkbox" checked={form.categorias.includes(c)} onChange={() => toggleCategoria(c)} />
              {c}
            </label>
          ))}
        </div>
        <div className="row" style={{ marginTop: 10 }}>
          <button className="primary" onClick={submit}>{editingId ? "Guardar cambios" : "Agregar profesor/a"}</button>
          {editingId && <button className="secondary" onClick={() => { setEditingId(null); setForm(emptyProfesorForm()); }}>Cancelar</button>}
        </div>
      </div>

      <div className="card">
        {profes.length === 0 && <div className="muted">No hay profesores registrados.</div>}
        {profes.map((p) => (
          <div key={p.id} className="list-item">
            <div>
              <b>{p.nombre}</b> {p.telefono && <span className="muted">— {p.telefono}</span>}
              <div className="muted">{(p.categorias || []).join(", ") || "Sin categorías asignadas"}</div>
            </div>
            <div>
              <button className="secondary" onClick={() => edit(p)}>Editar</button>{" "}
              <button className="secondary" onClick={() => remove(p.id)}>Eliminar</button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function Caja() {
  return (
    <div className="card">
      <div className="muted">Este módulo se va a ir completando con lo mismo que armamos en la versión de Claude (mensualidades, quiosco, profesores, etc). Por ahora es un placeholder para probar que el sistema de roles funciona.</div>
    </div>
  );
}
