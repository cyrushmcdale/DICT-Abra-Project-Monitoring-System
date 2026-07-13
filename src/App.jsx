import { useState, useEffect, useRef, Fragment } from "react";
import ExcelJS from 'exceljs';
import { supabase } from './supabaseClient';
import {
  LayoutDashboard,
  FolderOpen,
  Calendar,
  FileText,
  User,
  Users,
  Plus,
  Search,
  Eye,
  Pencil,
  Trash2,
  ArrowLeft,
  ChevronDown,
  ChevronRight,
  MapPin,
  Filter,
  Menu,
  X
} from 'lucide-react';

/* ─────────────────────────────────────────────────────────────
   DATA MODEL
   Backed by Supabase (Postgres). Schema + seed data: supabase/schema.sql.

     programs   -> table "programs"   (id, code, name, color, bg)
     events     -> table "events"     (id, program_id FK, name, start_date,
                                        end_date, location, personnel,
                                        description, status)
     employees  -> table "employees"  (id, name, role)
     notes      -> table "notes"      (program_id FK -> text)

   The in-memory `store` object mirrors these tables (events keyed by
   program_id to match the UI's existing access patterns). Every CRUD
   action updates local state immediately and fires the matching
   Supabase call. An event's completion is tracked solely by its
   `status` field (Target/Accomplished) — there is no separate numeric
   target/actual quantity.
───────────────────────────────────────────────────────────────*/

const COLOR_PRESETS = [
  { color: "#1a56db", bg: "#ebf5ff" },
  { color: "#0e9f6e", bg: "#f3faf7" },
  { color: "#d97706", bg: "#fffbeb" },
  { color: "#e02424", bg: "#fdf2f2" },
  { color: "#7e3af2", bg: "#f5f3ff" },
  { color: "#0891b2", bg: "#ecfeff" },
  { color: "#db2777", bg: "#fdf2f8" },
  { color: "#65a30d", bg: "#f7fee7" },
];

function slugify(str) {
  return (str || "").toLowerCase().trim().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "") || "program";
}

function makeUniqueId(base, existingIds) {
  let id = base, n = 1;
  while (existingIds.includes(id)) { id = `${base}_${n}`; n++; }
  return id;
}

function formatDateRange(ev) {
  const s = ev.startDate || ev.date || "";
  const e = ev.endDate || ev.date || s;
  if (!s) return "—";
  if (!e || e === s) return s;
  return `${s} → ${e}`;
}

// Whether an event's date range overlaps the selected reporting period.
// mode: "all" | "month" (month = "YYYY-MM") | "year" (year = number/string)
function eventInPeriod(ev, mode, month, year) {
  if (mode === "all") return true;
  const s = ev.startDate || ev.date || "";
  const e = ev.endDate || ev.startDate || ev.date || s;
  if (!s) return false;
  if (mode === "year") {
    const yStart = `${year}-01-01`;
    const yEnd = `${year}-12-31`;
    return e >= yStart && s <= yEnd;
  }
  if (mode === "month") {
    const [yy, mm] = month.split("-").map(Number);
    const lastDay = new Date(yy, mm, 0).getDate();
    const mStart = `${month}-01`;
    const mEnd = `${month}-${String(lastDay).padStart(2, "0")}`;
    return e >= mStart && s <= mEnd;
  }
  return true;
}

const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function PeriodFilter({ mode, setMode, month, setMonth, year, setYear }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
      <div style={{ display: "flex", background: "#f1f5f9", padding: 3, borderRadius: 8 }}>
        {[["all", "All Time"], ["month", "Month"], ["year", "Year"]].map(([val, label]) => (
          <button key={val} onClick={() => setMode(val)}
            style={{ padding: "6px 14px", border: "none", borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: "pointer", transition: "all 0.15s",
              background: mode === val ? "#0f172a" : "transparent", color: mode === val ? "#fff" : "#475569" }}>
            {label}
          </button>
        ))}
      </div>
      {mode === "month" && (
        <input type="month" value={month} onChange={e => setMonth(e.target.value)}
          style={{ padding: "7px 10px", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 12, outline: "none", color: "#334155" }} />
      )}
      {mode === "year" && (
        <input type="number" value={year} onChange={e => setYear(e.target.value)}
          style={{ width: 90, padding: "7px 10px", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 12, outline: "none", color: "#334155" }} />
      )}
    </div>
  );
}

function statusChip(p) {
  if (p === null) return <span style={chip("gray")}>No data</span>;
  if (p >= 90) return <span style={chip("green")}>Met</span>;
  if (p >= 65) return <span style={chip("amber")}>On track</span>;
  return <span style={chip("red")}>At risk</span>;
}

function chip(c) {
  const map = {
    green: { bg: "#def7ec", color: "#03543f" },
    amber: { bg: "#fdf6b2", color: "#723b13" },
    red:   { bg: "#fde8e8", color: "#9b1c1c" },
    gray:  { bg: "#f3f4f6", color: "#374151" },
  };
  const { bg, color } = map[c];
  return { background: bg, color, fontSize: 11, padding: "2px 8px", borderRadius: 10, fontWeight: 600, display: "inline-block", whiteSpace: "nowrap" };
}

function ProgressBar({ value, color }) {
  const clamped = Math.min(100, Math.max(0, value || 0));
  return (
    <div style={{ height: 8, background: "#f0f0f0", borderRadius: 4, overflow: "hidden", marginTop: 4 }}>
      <div style={{ width: `${clamped}%`, height: "100%", background: color, borderRadius: 4, transition: "width .4s" }} />
    </div>
  );
}

function Dot({ color, size = 8 }) {
  return <span style={{ display: "inline-block", width: size, height: size, borderRadius: "50%", background: color, flexShrink: 0 }} />;
}

function initialsOf(name) {
  return (name || "?").trim().split(/\s+/).slice(0, 2).map(w => w[0]).join("").toUpperCase();
}

// Multi-select employee dropdown: stores selection as a comma-joined name string
// so it stays compatible with the existing free-text `personnel` field.
function EmployeePicker({ employees, value, onChange, placeholder = "Select one or more employees..." }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);

  useEffect(() => {
    function onDocClick(e) {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, []);

  const selectedNames = (value || "").split(",").map(s => s.trim()).filter(Boolean);
  const selectedEmployees = (employees || []).filter(e => selectedNames.includes(e.name));

  const toggle = (name) => {
    const set = new Set(selectedNames);
    if (set.has(name)) set.delete(name); else set.add(name);
    onChange(Array.from(set).join(", "));
  };

  return (
    <div ref={ref} style={{ position: "relative", marginBottom: 14 }}>
      <button type="button" onClick={() => setOpen(o => !o)}
        style={{ width: "100%", display: "flex", alignItems: "center", gap: 10, padding: "8px 12px", border: "1px solid #d1d5db", borderRadius: 8, background: "#fff", cursor: "pointer", textAlign: "left", boxSizing: "border-box" }}>
        {selectedEmployees.length === 0 ? (
          <span style={{ fontSize: 13, color: "#94a3b8" }}>{placeholder}</span>
        ) : (
          <>
            <div style={{ display: "flex", flexShrink: 0 }}>
              {selectedEmployees.slice(0, 3).map((emp, i) => (
                <div key={emp.id} style={{ width: 26, height: 26, borderRadius: "50%", background: "#eff6ff", color: "#2563eb", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 700, border: "2px solid #fff", marginLeft: i === 0 ? 0 : -8 }}>
                  {initialsOf(emp.name)}
                </div>
              ))}
            </div>
            <div style={{ fontSize: 13, fontWeight: 600, color: "#1e293b", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
              {selectedEmployees.length === 1 ? selectedEmployees[0].name : `${selectedEmployees[0].name} +${selectedEmployees.length - 1} more`}
            </div>
          </>
        )}
        <ChevronDown size={14} style={{ marginLeft: "auto", color: "#94a3b8", flexShrink: 0 }} />
      </button>

      {open && (
        <div style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, background: "#fff", border: "1px solid #e2e8f0", borderRadius: 10, boxShadow: "0 10px 30px rgba(0,0,0,.12)", zIndex: 20, maxHeight: 240, overflowY: "auto" }}>
          {(employees || []).length === 0 && (
            <div style={{ padding: 12, fontSize: 12, color: "#94a3b8", textAlign: "center" }}>No employees yet. Add one under Employees.</div>
          )}
          {(employees || []).map(emp => {
            const checked = selectedNames.includes(emp.name);
            return (
              <button key={emp.id} type="button" onClick={() => toggle(emp.name)}
                style={{ width: "100%", display: "flex", alignItems: "center", gap: 10, padding: "8px 12px", border: "none", background: checked ? "#eff6ff" : "#fff", cursor: "pointer", textAlign: "left" }}>
                <div style={{ width: 28, height: 28, borderRadius: "50%", background: "#eff6ff", color: "#2563eb", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 700, flexShrink: 0 }}>
                  {initialsOf(emp.name)}
                </div>
                <div style={{ lineHeight: 1.25, flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: "#1e293b" }}>{emp.name}</div>
                  <div style={{ fontSize: 11, color: "#94a3b8" }}>{emp.role || "—"}</div>
                </div>
                {checked && <span style={{ color: "#2563eb", fontWeight: 700, flexShrink: 0 }}>✓</span>}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

const inputStyle = { width: "100%", padding: "9px 12px", border: "1px solid #d1d5db", borderRadius: 8, fontSize: 13, outline: "none", marginBottom: 14, boxSizing: "border-box" };
const labelStyle = { fontSize: 12, fontWeight: 600, color: "#374151", display: "block", marginBottom: 4 };
const overlayStyle = { position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 };
const cancelBtnStyle = { padding: "8px 18px", border: "1px solid #e2e8f0", borderRadius: 8, background: "#fff", fontSize: 13, cursor: "pointer", color: "#475569" };

export default function App() {
  // Authentication States
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [authChecked, setAuthChecked] = useState(false);
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [loginLoading, setLoginLoading] = useState(false);

  // Application Dashboard States
  const [store, setStore] = useState(null);
  const [dataError, setDataError] = useState("");
  const [view, setView] = useState("dashboard");
  const [activeProgram, setActiveProgram] = useState("");
  const [toast, setToast] = useState(null);

  // Event (formerly "indicator") add-form states (used by the Add Event modal)
  const [newEventName, setNewEventName] = useState("");
  const [newEventStartDate, setNewEventStartDate] = useState("");
  const [newEventEndDate, setNewEventEndDate] = useState("");
  const [newEventLocation, setNewEventLocation] = useState("");
  const [newEventPersonnel, setNewEventPersonnel] = useState("");
  const [newEventDescription, setNewEventDescription] = useState("");
  const [newEventStatus, setNewEventStatus] = useState("Target");
  const [deleteEventConfirm, setDeleteEventConfirm] = useState(null);

  // Edit Event modal states
  const [editEventModalOpen, setEditEventModalOpen] = useState(false);
  const [editEventForm, setEditEventForm] = useState(null);

  // Events tab search/filter/sort states
  const [eventsSearch, setEventsSearch] = useState("");
  const [eventsDateFrom, setEventsDateFrom] = useState("");
  const [eventsDateTo, setEventsDateTo] = useState("");
  const [eventsSort, setEventsSort] = useState("date_asc");

  // Shared reporting-period filter (All Time / Month / Year) — drives Dashboard,
  // Programs, Reports, and Events completion metrics consistently.
  const [periodMode, setPeriodMode] = useState("all");
  const [periodMonth, setPeriodMonth] = useState(() => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  });
  const [periodYear, setPeriodYear] = useState(() => new Date().getFullYear());

  // Employees CRUD states
  const [employeeModalOpen, setEmployeeModalOpen] = useState(false);
  const [employeeModalMode, setEmployeeModalMode] = useState("add");
  const [employeeForm, setEmployeeForm] = useState({ id: null, name: "", role: "" });
  const [deleteEmployeeConfirm, setDeleteEmployeeConfirm] = useState(null);

  // Events sidebar dropdown (Events / + Add Event sub-tabs)
  const [eventsMenuOpen, setEventsMenuOpen] = useState(false);

  // Notes states
  const [notesOpen, setNotesOpen] = useState(false);
  const [noteText, setNoteText] = useState("");

  // Programs list/search + CRUD modal states
  const [programSearch, setProgramSearch] = useState("");
  const [programModalOpen, setProgramModalOpen] = useState(false);
  const [programModalMode, setProgramModalMode] = useState("add");
  const [programForm, setProgramForm] = useState({ id: null, code: "", name: "", colorIdx: 0 });
  const [expandedEventDetails, setExpandedEventDetails] = useState(null);

  // Responsive layout: collapsible sidebar on narrow (phone) viewports
  const [isMobile, setIsMobile] = useState(() => window.innerWidth <= 900);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  useEffect(() => {
    const mq = window.matchMedia("(max-width: 900px)");
    const handler = (e) => setIsMobile(e.matches);
    mq.addEventListener("change", handler);
    return () => mq.removeEventListener("change", handler);
  }, []);

  // Calendar Engine Management States
  const [currentCalendarYear, setCurrentCalendarYear] = useState(() => new Date().getFullYear());
  const [currentCalendarMonth, setCurrentCalendarMonth] = useState(() => new Date().getMonth());
  const [selectedCalendarDate, setSelectedCalendarDate] = useState(() => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  });

  // Track the Supabase session so isAuthenticated always reflects reality
  // (e.g. if the session expires in another tab).
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setIsAuthenticated(!!session);
      setAuthChecked(true);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      setIsAuthenticated(!!session);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  // Load all data from Supabase once signed in.
  useEffect(() => {
    if (!isAuthenticated) return;
    let cancelled = false;
    (async () => {
      const [programsRes, eventsRes, employeesRes, notesRes] = await Promise.all([
        supabase.from("programs").select("*").order("created_at"),
        supabase.from("events").select("*").order("start_date"),
        supabase.from("employees").select("*").order("created_at"),
        supabase.from("notes").select("*"),
      ]);
      if (cancelled) return;
      if (programsRes.error || eventsRes.error || employeesRes.error || notesRes.error) {
        setDataError("Couldn't load data from the database. Check your connection and try again.");
        return;
      }
      const programs = (programsRes.data || []).map(p => ({ id: p.id, code: p.code, name: p.name, color: p.color, bg: p.bg }));
      const events = {};
      (eventsRes.data || []).forEach(ev => {
        const mapped = {
          id: ev.id, name: ev.name, startDate: ev.start_date || "", endDate: ev.end_date || ev.start_date || "",
          location: ev.location || "", personnel: ev.personnel || "", description: ev.description || "", status: ev.status || "Target",
        };
        (events[ev.program_id] = events[ev.program_id] || []).push(mapped);
      });
      const employees = (employeesRes.data || []).map(e => ({ id: e.id, name: e.name, role: e.role }));
      const notes = {};
      (notesRes.data || []).forEach(n => { notes[`${n.program_id}_notes`] = n.text || ""; });

      setStore({ programs, events, employees, notes, schedule: {} });
      setActiveProgram(prev => prev || programs[0]?.id || "");
      setDataError("");
    })();
    return () => { cancelled = true; };
  }, [isAuthenticated]);

  const showToast = (msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 2800);
  };

  // Auth Handlers
  const handleLogin = async (e) => {
    e.preventDefault();
    setLoginError("");
    setLoginLoading(true);
    const { error } = await supabase.auth.signInWithPassword({ email: username.trim(), password });
    setLoginLoading(false);
    if (error) {
      setLoginError(error.message || "Invalid email or password.");
      return;
    }
    setUsername("");
    setPassword("");
    setTimeout(() => showToast("Welcome back, Administrator!"), 300);
  };

  const handleLogout = async () => {
    if (!window.confirm("Are you sure you want to log out?")) return;
    await supabase.auth.signOut();
    setStore(null);
    setView("dashboard");
  };

  // Navigate and, on mobile, close the slide-in nav drawer.
  const goTo = (v) => {
    setView(v);
    setMobileNavOpen(false);
  };

  // ── EVENT (per-program) CRUD ──
  const addEvent = async () => {
    if (!newEventName.trim() || !activeProgram) return;
    const id = `${activeProgram}_custom_${Date.now()}`;
    const newEvent = {
      id,
      name: newEventName.trim(),
      startDate: newEventStartDate || "",
      endDate: newEventEndDate || newEventStartDate || "",
      location: newEventLocation.trim(),
      personnel: newEventPersonnel.trim() || "Provincial Project Officer",
      description: newEventDescription.trim(),
      status: newEventStatus || "Target",
    };
    const { error } = await supabase.from("events").insert({
      id: newEvent.id,
      program_id: activeProgram,
      name: newEvent.name,
      start_date: newEvent.startDate || null,
      end_date: newEvent.endDate || null,
      location: newEvent.location,
      personnel: newEvent.personnel,
      description: newEvent.description,
      status: newEvent.status,
    });
    if (error) { showToast("Failed to save event.", "error"); return; }
    setStore(prev => ({
      ...prev,
      events: { ...prev.events, [activeProgram]: [...(prev.events[activeProgram] || []), newEvent] },
    }));
    setNewEventName(""); setNewEventStartDate(""); setNewEventEndDate(""); setNewEventLocation(""); setNewEventPersonnel(""); setNewEventDescription(""); setNewEventStatus("Target");
    showToast("Event added.");
  };

  // Field-name mapping between the in-memory event shape and the "events" table columns.
  const EVENT_FIELD_TO_COLUMN = { startDate: "start_date", endDate: "end_date" };

  const updateEventField = async (progId, eventId, field, value) => {
    setStore(prev => ({
      ...prev,
      events: {
        ...prev.events,
        [progId]: (prev.events[progId] || []).map(ev => ev.id === eventId ? { ...ev, [field]: value } : ev),
      },
    }));
    const column = EVENT_FIELD_TO_COLUMN[field] || field;
    const { error } = await supabase.from("events").update({ [column]: value || null }).eq("id", eventId);
    if (error) showToast("Failed to save change.", "error");
  };

  const openEditEventModal = (progId, ev) => {
    setEditEventForm({
      progId,
      id: ev.id,
      name: ev.name || "",
      startDate: ev.startDate || ev.date || "",
      endDate: ev.endDate || ev.date || "",
      location: ev.location || "",
      personnel: ev.personnel || "",
      description: ev.description || "",
      status: ev.status || "Target",
    });
    setEditEventModalOpen(true);
  };

  const saveEditEvent = async () => {
    const f = editEventForm;
    if (!f || !f.name.trim()) return;
    const name = f.name.trim();
    const startDate = f.startDate || "";
    const endDate = f.endDate || f.startDate || "";
    const { error } = await supabase.from("events").update({
      name,
      start_date: startDate || null,
      end_date: endDate || null,
      location: f.location,
      personnel: f.personnel,
      description: f.description,
      status: f.status,
    }).eq("id", f.id);
    if (error) { showToast("Failed to update event.", "error"); return; }
    setStore(prev => ({
      ...prev,
      events: {
        ...prev.events,
        [f.progId]: (prev.events[f.progId] || []).map(ev => ev.id === f.id
          ? { ...ev, name, startDate, endDate, location: f.location, personnel: f.personnel, description: f.description, status: f.status }
          : ev),
      },
    }));
    setEditEventModalOpen(false);
    showToast("Event updated.");
  };

  const toggleEventStatus = async (progId, eventId, currentStatus) => {
    await updateEventField(progId, eventId, "status", currentStatus === "Accomplished" ? "Target" : "Accomplished");
    showToast(currentStatus === "Accomplished" ? "Marked as Target." : "Marked as Accomplished!");
  };

  const deleteEvent = async (progId, eventId) => {
    const { error } = await supabase.from("events").delete().eq("id", eventId);
    if (error) { showToast("Failed to delete event.", "error"); return; }
    setStore(prev => ({
      ...prev,
      events: { ...prev.events, [progId]: (prev.events[progId] || []).filter(e => e.id !== eventId) },
    }));
    setDeleteEventConfirm(null);
    showToast("Event removed.", "info");
  };

  // ── PROGRAM CRUD ──
  const openAddProgramModal = () => {
    setProgramModalMode("add");
    setProgramForm({ id: null, code: "", name: "", colorIdx: store.programs.length % COLOR_PRESETS.length });
    setProgramModalOpen(true);
  };

  const openEditProgramModal = (p) => {
    const colorIdx = COLOR_PRESETS.findIndex(c => c.color === p.color);
    setProgramModalMode("edit");
    setProgramForm({ id: p.id, code: p.code, name: p.name, colorIdx: colorIdx >= 0 ? colorIdx : 0 });
    setProgramModalOpen(true);
  };

  const saveProgram = async () => {
    if (!programForm.code.trim() || !programForm.name.trim()) return;
    const preset = COLOR_PRESETS[programForm.colorIdx] || COLOR_PRESETS[0];

    if (programModalMode === "add") {
      const baseId = slugify(programForm.code);
      const id = makeUniqueId(baseId, store.programs.map(p => p.id));
      const newProgram = { id, code: programForm.code.trim().toUpperCase(), name: programForm.name.trim(), color: preset.color, bg: preset.bg };
      const { error } = await supabase.from("programs").insert(newProgram);
      if (error) { showToast("Failed to add program.", "error"); return; }
      setStore(prev => ({
        ...prev,
        programs: [...prev.programs, newProgram],
        events: { ...prev.events, [id]: [] },
      }));
      showToast("Program added.");
    } else {
      const updates = { code: programForm.code.trim().toUpperCase(), name: programForm.name.trim(), color: preset.color, bg: preset.bg };
      const { error } = await supabase.from("programs").update(updates).eq("id", programForm.id);
      if (error) { showToast("Failed to update program.", "error"); return; }
      setStore(prev => ({
        ...prev,
        programs: prev.programs.map(p => p.id === programForm.id ? { ...p, ...updates } : p),
      }));
      showToast("Program updated.");
    }
    setProgramModalOpen(false);
  };

  // ── EMPLOYEE CRUD ──
  const openAddEmployeeModal = () => {
    setEmployeeModalMode("add");
    setEmployeeForm({ id: null, name: "", role: "" });
    setEmployeeModalOpen(true);
  };

  const openEditEmployeeModal = (emp) => {
    setEmployeeModalMode("edit");
    setEmployeeForm({ id: emp.id, name: emp.name, role: emp.role });
    setEmployeeModalOpen(true);
  };

  const saveEmployee = async () => {
    if (!employeeForm.name.trim()) return;
    if (employeeModalMode === "add") {
      const id = makeUniqueId(slugify(employeeForm.name), (store.employees || []).map(e => e.id));
      const newEmployee = { id, name: employeeForm.name.trim(), role: employeeForm.role.trim() };
      const { error } = await supabase.from("employees").insert(newEmployee);
      if (error) { showToast("Failed to add employee.", "error"); return; }
      setStore(prev => ({ ...prev, employees: [...(prev.employees || []), newEmployee] }));
      showToast("Employee added.");
    } else {
      const updates = { name: employeeForm.name.trim(), role: employeeForm.role.trim() };
      const { error } = await supabase.from("employees").update(updates).eq("id", employeeForm.id);
      if (error) { showToast("Failed to update employee.", "error"); return; }
      setStore(prev => ({
        ...prev,
        employees: (prev.employees || []).map(e => e.id === employeeForm.id ? { ...e, ...updates } : e),
      }));
      showToast("Employee updated.");
    }
    setEmployeeModalOpen(false);
  };

  const deleteEmployee = async (id) => {
    const { error } = await supabase.from("employees").delete().eq("id", id);
    if (error) { showToast("Failed to remove employee.", "error"); return; }
    setStore(prev => ({ ...prev, employees: (prev.employees || []).filter(e => e.id !== id) }));
    setDeleteEmployeeConfirm(null);
    showToast("Employee removed.", "info");
  };

  const saveNote = async () => {
    const { error } = await supabase.from("notes").upsert({ program_id: activeProgram, text: noteText, updated_at: new Date().toISOString() });
    if (error) { showToast("Failed to save remarks.", "error"); return; }
    setStore(prev => ({ ...prev, notes: { ...prev.notes, [noteKey]: noteText } }));
    setNotesOpen(false);
    showToast("Remarks saved.");
  };

  const exportToExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "DICT Abra Project Monitoring System";
    workbook.created = new Date();

    const tableColumns = [
      { name: "Event ID" },
      { name: "Program Event" },
      { name: "Start Date" },
      { name: "End Date" },
      { name: "Location" },
      { name: "Assigned Employee/s" },
      { name: "Event Status" },
    ];

    store.programs.forEach((p, pIdx) => {
      const pevs = (store.events[p.id] || []).filter(ev => eventInPeriod(ev, periodMode, periodMonth, periodYear));
      const sheetName = (p.code || `Program${pIdx + 1}`).replace(/[\\/?*[\]:]/g, "").slice(0, 31) || `Program${pIdx + 1}`;
      const worksheet = workbook.addWorksheet(sheetName);
      worksheet.columns = [{ width: 20 }, { width: 34 }, { width: 14 }, { width: 14 }, { width: 30 }, { width: 30 }, { width: 16 }];

      if (pevs.length > 0) {
        const tableRows = pevs.map(ev => [
          ev.id,
          ev.name,
          ev.startDate || ev.date || "",
          ev.endDate || ev.startDate || ev.date || "",
          ev.location || "",
          ev.personnel || "",
          ev.status || "Target",
        ]);
        worksheet.addTable({
          name: `Table_${slugify(sheetName)}_${pIdx}`.replace(/[^A-Za-z0-9_]/g, "_"),
          ref: "A1",
          headerRow: true,
          style: { theme: "TableStyleMedium9", showRowStripes: true },
          columns: tableColumns,
          rows: tableRows,
        });
      } else {
        const headerRow = worksheet.addRow(tableColumns.map(c => c.name));
        headerRow.eachCell(cell => {
          cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
        });
        worksheet.addRow(["No events for the selected period."]);
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const stamp = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = `DICT_Abra_Consolidated_Report_${stamp}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast("Excel workbook exported with formatted tables!");
  };

  const getOverall = (progId) => {
    const evs = (store.events[progId] || []).filter(ev => eventInPeriod(ev, periodMode, periodMonth, periodYear));
    if (!evs.length) return null;
    const accomplished = evs.filter(ev => ev.status === "Accomplished").length;
    return Math.round((accomplished / evs.length) * 100);
  };

  // Wait for the initial Supabase session check before deciding which screen to show,
  // so a logged-in user doesn't flash the login screen on refresh.
  if (!authChecked) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#0f172a", color: "#94a3b8", fontFamily: "'IBM Plex Sans', sans-serif", fontSize: 13 }}>
        Loading…
      </div>
    );
  }

  // Login Screen UI Render Guard
  if (!isAuthenticated) {
    return (
      <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", minHeight: "100vh", background: "#0f172a", display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
        <div style={{ background: "#1e293b", padding: "40px 32px", borderRadius: 16, width: "100%", maxWidth: 400, boxShadow: "0 25px 50px -12px rgba(0,0,0,0.5)", border: "1px solid #334155" }}>

          <div style={{ textAlign: "center", marginBottom: 28 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: "#38bdf8", letterSpacing: 2, marginBottom: 4 }}>DICT ABRA PROVINCE</div>
            <h1 style={{ fontSize: 20, fontWeight: 700, color: "#f8fafc", margin: 0, lineHeight: 1.3 }}>Project Monitoring System</h1>
            <p style={{ fontSize: 12, color: "#94a3b8", marginTop: 6 }}>Sign in to encode data and manage metrics</p>
          </div>

          <form onSubmit={handleLogin}>
            {loginError && (
              <div style={{ background: "#7f1d1d", border: "1px solid #f87171", color: "#fca5a5", fontSize: 12, padding: "10px 12px", borderRadius: 8, marginBottom: 16, fontWeight: 500 }}>
                ⚠️ {loginError}
              </div>
            )}

            <div style={{ marginBottom: 16 }}>
              <label style={{ display: "block", fontSize: 12, color: "#cbd5e1", fontWeight: 600, marginBottom: 6 }}>Email</label>
              <input type="email" value={username} onChange={e => setUsername(e.target.value)} required placeholder="you@office.gov.ph"
                style={{ width: "100%", padding: "11px 14px", background: "#0f172a", border: "1px solid #475569", borderRadius: 8, color: "#f8fafc", fontSize: 14, outline: "none", boxSizing: "border-box" }} />
            </div>

            <div style={{ marginBottom: 24 }}>
              <label style={{ display: "block", fontSize: 12, color: "#cbd5e1", fontWeight: 600, marginBottom: 6 }}>Password</label>
              <input type="password" value={password} onChange={e => setPassword(e.target.value)} required placeholder="••••••••"
                style={{ width: "100%", padding: "11px 14px", background: "#0f172a", border: "1px solid #475569", borderRadius: 8, color: "#f8fafc", fontSize: 14, outline: "none", boxSizing: "border-box" }} />
            </div>

            <button type="submit" disabled={loginLoading}
              style={{ width: "100%", padding: "12px", background: "#2563eb", color: "#fff", border: "none", borderRadius: 8, fontSize: 14, fontWeight: 600, cursor: loginLoading ? "not-allowed" : "pointer", opacity: loginLoading ? 0.7 : 1, transition: "background 0.2s" }}>
              {loginLoading ? "Signing in…" : "Sign In"}
            </button>
          </form>

          <div style={{ textAlign: "center", marginTop: 24, fontSize: 11, color: "#64748b", borderTop: "1px solid #334155", paddingTop: 16 }}>
            Use your DICT Abra office account credentials.
          </div>
        </div>
      </div>
    );
  }

  // Data Loading Guard — the Supabase fetch effect runs after isAuthenticated flips true
  if (!store) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#f8fafc", fontFamily: "'IBM Plex Sans', 'Segoe UI', sans-serif", color: "#64748b", fontSize: 14 }}>
        {dataError ? (
          <div style={{ textAlign: "center" }}>
            <p style={{ color: "#e02424", fontWeight: 600, marginBottom: 12 }}>{dataError}</p>
            <button onClick={() => window.location.reload()}
              style={{ padding: "8px 18px", background: "#0f172a", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, cursor: "pointer" }}>
              Retry
            </button>
          </div>
        ) : "Loading data…"}
      </div>
    );
  }

  const prog = store.programs.find(p => p.id === activeProgram) || null;
  const events = prog ? (store.events[prog.id] || []) : [];

  const overallAll = (() => {
    const vals = store.programs.map(p => getOverall(p.id)).filter(v => v !== null);
    if (!vals.length) return null;
    return Math.round(vals.reduce((a, b) => a + b, 0) / vals.length);
  })();

  const noteKey = `${activeProgram}_notes`;

  const filteredPrograms = store.programs.filter(p => {
    const term = programSearch.trim().toLowerCase();
    if (!term) return true;
    return p.code.toLowerCase().includes(term) || p.name.toLowerCase().includes(term);
  });

  const totalEvents = store.programs.reduce((s, p) => s + (store.events[p.id] || []).length, 0);

  // ── Program-level event counts by status (Target = pending, Accomplished = done) ──
  let maxVal = 5;
  const chartData = store.programs.map(p => {
    const evs = (store.events[p.id] || []).filter(ev => eventInPeriod(ev, periodMode, periodMonth, periodYear));
    const target = evs.filter(ev => ev.status !== "Accomplished").length;
    const actual = evs.filter(ev => ev.status === "Accomplished").length;
    if (target > maxVal) maxVal = target;
    if (actual > maxVal) maxVal = actual;
    return { id: p.id, code: p.code, color: p.color, target, actual };
  });
  const yAxisMax = Math.ceil((maxVal * 1.15) / 5) * 5;

  // ── Single source of truth for events across Dashboard, Programs tab, and Calendar ──
  const allEventsFlat = store.programs.flatMap(p =>
    (store.events[p.id] || []).map(ev => ({ ...ev, progId: p.id }))
  );
  // All pending (Target) events, all time — drives the "Upcoming Events" summary card.
  const upcomingEvents = allEventsFlat
    .filter(ev => (ev.startDate || ev.date) && ev.status !== "Accomplished")
    .sort((a, b) => (a.startDate > b.startDate ? 1 : a.startDate < b.startDate ? -1 : 0));

  // Same list narrowed to the selected reporting period, grouped by month for the Dashboard table.
  const upcomingEventsInPeriod = upcomingEvents.filter(ev => eventInPeriod(ev, periodMode, periodMonth, periodYear));
  const upcomingEventsByMonth = (() => {
    const groups = {};
    upcomingEventsInPeriod.forEach(ev => {
      const key = (ev.startDate || ev.date || "").slice(0, 7);
      if (!key) return;
      (groups[key] = groups[key] || []).push(ev);
    });
    return Object.keys(groups).sort().map(key => {
      const [y, m] = key.split("-").map(Number);
      return { key, label: `${MONTH_NAMES[m - 1]} ${y}`, events: groups[key] };
    });
  })();

  const eventsOnDate = (dateStr) => allEventsFlat.filter(ev => {
    const s = ev.startDate || ev.date || "";
    const e = ev.endDate || ev.startDate || ev.date || "";
    return s && dateStr >= s && dateStr <= e;
  });

  // ── Events tab: all events across all programs, regardless of status, with search/date-range/sort ──
  const filteredAllEvents = allEventsFlat
    .filter(ev => eventInPeriod(ev, periodMode, periodMonth, periodYear))
    .filter(ev => {
      const term = eventsSearch.trim().toLowerCase();
      if (!term) return true;
      return ev.name.toLowerCase().includes(term)
        || (ev.personnel || "").toLowerCase().includes(term)
        || (ev.location || "").toLowerCase().includes(term);
    })
    .filter(ev => {
      const s = ev.startDate || ev.date || "";
      const e = ev.endDate || ev.startDate || ev.date || "";
      if (eventsDateFrom && (!e || e < eventsDateFrom)) return false;
      if (eventsDateTo && (!s || s > eventsDateTo)) return false;
      return true;
    })
    .sort((a, b) => {
      if (eventsSort === "alpha") return a.name.localeCompare(b.name);
      const aStart = a.startDate || a.date || "";
      const bStart = b.startDate || b.date || "";
      if (eventsSort === "date_desc") return bStart.localeCompare(aStart);
      return aStart.localeCompare(bStart);
    });

  // Building Grid Mathematics Matrix for Calendar rendering
  const daysInMonth = new Date(currentCalendarYear, currentCalendarMonth + 1, 0).getDate();
  const firstDayIndex = new Date(currentCalendarYear, currentCalendarMonth, 1).getDay();

  const calendarDays = [];
  for (let i = 0; i < firstDayIndex; i++) calendarDays.push(null);
  for (let d = 1; d <= daysInMonth; d++) calendarDays.push(d);

  const changeMonth = (direction) => {
    if (direction === "next") {
      if (currentCalendarMonth === 11) {
        setCurrentCalendarMonth(0);
        setCurrentCalendarYear(p => p + 1);
      } else {
        setCurrentCalendarMonth(p => p + 1);
      }
    } else {
      if (currentCalendarMonth === 0) {
        setCurrentCalendarMonth(11);
        setCurrentCalendarYear(p => p - 1);
      } else {
        setCurrentCalendarMonth(p => p - 1);
      }
    }
  };

  const inEventsSection = view === "events" || view === "addEvent";
  const showEventsSubmenu = eventsMenuOpen || inEventsSection;

  return (
    <div style={{ fontFamily: "'IBM Plex Sans', 'Segoe UI', sans-serif", minHeight: "100vh", background: "#f8fafc", color: "#1e293b" }}>

      {/* SIDEBAR NAVIGATION PANEL */}
      <div style={{ display: "flex", minHeight: "100vh" }}>
        {isMobile && mobileNavOpen && (
          <div onClick={() => setMobileNavOpen(false)}
            style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,.5)", zIndex: 1050 }} />
        )}
        <nav style={{ width: isMobile ? "min(260px, 85vw)" : 260, background: "#fff", display: "flex", flexDirection: "column", padding: "16px", flexShrink: 0,
          position: isMobile ? "fixed" : "sticky", top: 0, left: 0, height: "100vh", borderRight: "1px solid #f1f5f9",
          zIndex: 1100, transform: isMobile ? (mobileNavOpen ? "translateX(0)" : "translateX(-100%)") : "none",
          transition: "transform .25s ease", boxShadow: isMobile && mobileNavOpen ? "4px 0 24px rgba(0,0,0,.2)" : "none" }}>

          {/* BRANDING SECTION (Logo left, stacked text right) */}
          <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "0 8px 16px", marginBottom: 24, borderBottom: "1px solid #f1f5f9" }}>
            <img
              src="/dict-logo.png"
              alt="DICT Logo"
              style={{ width: 40, height: 40, objectFit: "contain", flexShrink: 0 }}
            />
            <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.2, flex: 1 }}>
              <span style={{ fontSize: 16, fontWeight: 700, tracking: "0.025em", color: "#1e293b" }}>DICT ABRA</span>
              <span style={{ fontSize: 11, fontWeight: 500, color: "#94a3b8", textTransform: "uppercase" }}>
                Project Monitoring System
              </span>
            </div>
            {isMobile && (
              <button onClick={() => setMobileNavOpen(false)} title="Close menu"
                style={{ border: "none", background: "none", color: "#94a3b8", cursor: "pointer", padding: 4, display: "flex", flexShrink: 0 }}>
                <X size={20} />
              </button>
            )}
          </div>

          <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 2 }}>
            <div style={{ fontSize: 11, color: "#94a3b8", fontWeight: 700, letterSpacing: 1, paddingLeft: 12, marginBottom: 8, textTransform: "uppercase" }}>MENU</div>

            {/* 1. Dashboard View Link */}
            <button onClick={() => goTo("dashboard")}
              style={{ display: "flex", alignItems: "center", gap: 12, width: "100%", padding: "10px 12px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 14, transition: "all 0.15s",
                background: view === "dashboard" ? "#eff6ff" : "transparent",
                color: view === "dashboard" ? "#2563eb" : "#64748b", fontWeight: view === "dashboard" ? 600 : 500, marginBottom: 4 }}>
              <LayoutDashboard size={18} strokeWidth={view === "dashboard" ? 2.5 : 2} />
              Dashboard
            </button>

            {/* 2. Programs View Link */}
            <button onClick={() => goTo("programs")}
              style={{ display: "flex", alignItems: "center", gap: 12, width: "100%", padding: "10px 12px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 14, transition: "all 0.15s",
                background: (view === "programs" || view === "programDetail") ? "#eff6ff" : "transparent",
                color: (view === "programs" || view === "programDetail") ? "#2563eb" : "#64748b", fontWeight: (view === "programs" || view === "programDetail") ? 600 : 500, marginBottom: 4 }}>
              <FolderOpen size={18} strokeWidth={(view === "programs" || view === "programDetail") ? 2.5 : 2} />
              Programs
            </button>

            {/* 3. Events Dropdown (Events / + Add Event sub-tabs) */}
            <div style={{ marginBottom: 4 }}>
              <button onClick={() => setEventsMenuOpen(o => !o)}
                style={{ display: "flex", alignItems: "center", gap: 12, width: "100%", padding: "10px 12px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 14, transition: "all 0.15s",
                  background: inEventsSection ? "#eff6ff" : "transparent",
                  color: inEventsSection ? "#2563eb" : "#64748b", fontWeight: inEventsSection ? 600 : 500 }}>
                <Calendar size={18} strokeWidth={inEventsSection ? 2.5 : 2} />
                <span style={{ flex: 1, textAlign: "left" }}>Events</span>
                {showEventsSubmenu ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
              </button>

              {showEventsSubmenu && (
                <div style={{ display: "flex", flexDirection: "column", gap: 2, marginTop: 2, paddingLeft: 14 }}>
                  <button onClick={() => goTo("events")}
                    style={{ display: "flex", alignItems: "center", width: "100%", padding: "8px 12px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 13, transition: "all 0.15s",
                      background: view === "events" ? "#eff6ff" : "transparent",
                      color: view === "events" ? "#2563eb" : "#64748b",
                      fontWeight: view === "events" ? 600 : 500 }}>
                    Events
                  </button>
                  <button onClick={() => goTo("addEvent")}
                    style={{ display: "flex", alignItems: "center", width: "100%", padding: "8px 12px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 13, transition: "all 0.15s",
                      background: view === "addEvent" ? "#eff6ff" : "transparent",
                      color: view === "addEvent" ? "#2563eb" : "#64748b",
                      fontWeight: view === "addEvent" ? 600 : 500 }}>
                    + Add Event
                  </button>
                </div>
              )}
            </div>

            {/* 4. Employees View Link */}
            <button onClick={() => goTo("employees")}
              style={{ display: "flex", alignItems: "center", gap: 12, width: "100%", padding: "10px 12px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 14, transition: "all 0.15s",
                background: view === "employees" ? "#eff6ff" : "transparent",
                color: view === "employees" ? "#2563eb" : "#64748b", fontWeight: view === "employees" ? 600 : 500, marginBottom: 4 }}>
              <Users size={18} strokeWidth={view === "employees" ? 2.5 : 2} />
              Employees
            </button>

            {/* 5. Schedule View Link */}
            <button onClick={() => goTo("schedule")}
              style={{ display: "flex", alignItems: "center", gap: 12, width: "100%", padding: "10px 12px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 14, transition: "all 0.15s",
                background: view === "schedule" ? "#eff6ff" : "transparent",
                color: view === "schedule" ? "#2563eb" : "#64748b", fontWeight: view === "schedule" ? 600 : 500, marginBottom: 4 }}>
              <Calendar size={18} strokeWidth={view === "schedule" ? 2.5 : 2} />
              Schedule
            </button>

            {/* 6. Full Report Link */}
            <button onClick={() => goTo("report")}
              style={{ display: "flex", alignItems: "center", gap: 12, width: "100%", padding: "10px 12px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 14, transition: "all 0.15s",
                background: view === "report" ? "#eff6ff" : "transparent",
                color: view === "report" ? "#2563eb" : "#64748b", fontWeight: view === "report" ? 600 : 500, marginBottom: 4 }}>
              <FileText size={18} strokeWidth={view === "report" ? 2.5 : 2} />
              Reports
            </button>

            {/* GENERAL SECTION BOTTOM */}
            <div style={{ borderTop: "1px solid #f1f5f9", paddingTop: 16, marginTop: "auto" }}>
              <div style={{ fontSize: 11, color: "#94a3b8", fontWeight: 700, letterSpacing: 1, paddingLeft: 12, marginBottom: 8, textTransform: "uppercase" }}>GENERAL</div>

              <button onClick={exportToExcel}
                style={{ display: "flex", alignItems: "center", gap: 12, width: "100%", padding: "10px 12px", borderRadius: 8, border: "none", background: "transparent", color: "#64748b", fontSize: 14, cursor: "pointer", marginBottom: 4, fontWeight: 500 }}>
                <span style={{ fontSize: 16, lineHeight: 1 }}>📥</span> Export XLSX
              </button>
              <button onClick={handleLogout}
                style={{ display: "flex", alignItems: "center", gap: 12, width: "100%", padding: "10px 12px", borderRadius: 8, border: "none", background: "transparent", color: "#64748b", fontSize: 14, cursor: "pointer", fontWeight: 500 }}>
                <User size={18} />
                Log Out
              </button>
            </div>
          </div>
        </nav>

        {/* WORKSPACE APP PANELS RENDERER */}
        <main style={{ flex: 1, padding: isMobile ? "16px" : "28px 32px", maxWidth: isMobile ? "100%" : "calc(100vw - 260px)", overflowX: "hidden" }}>

          {isMobile && (
            <button onClick={() => setMobileNavOpen(true)} title="Open menu"
              style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 14px", background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 13, fontWeight: 600, color: "#334155", cursor: "pointer", marginBottom: 16 }}>
              <Menu size={18} /> Menu
            </button>
          )}

          {/* ── SECTION A: DASHBOARD VIEW ── */}
          {view === "dashboard" && (
            <div>
              <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 24 }}>
                <div>
                  <h1 style={{ fontSize: 22, fontWeight: 700, margin: 0 }}>Dashboard</h1>
                  <p style={{ fontSize: 13, color: "#64748b", margin: "4px 0 0" }}>DICT Abra Province — Consolidated Event Status Overview</p>
                </div>
                <PeriodFilter mode={periodMode} setMode={setPeriodMode} month={periodMonth} setMonth={setPeriodMonth} year={periodYear} setYear={setPeriodYear} />
              </div>

              {/* Summary cards matrix */}
              <div style={{ display: "grid", gridTemplateColumns: isMobile ? "repeat(2, 1fr)" : "repeat(5, 1fr)", gap: 14, marginBottom: 28 }}>
                {[
                  { label: "Upcoming Events", val: `${upcomingEvents.length}`, sub: "All pending (Target) events", accent: "#2563eb" },
                  { label: "Overall Completion Rate", val: overallAll !== null ? `${overallAll}%` : "—", sub: "All programs combined", accent: overallAll !== null ? (overallAll >= 75 ? "#0e9f6e" : "#d97706") : "#64748b" },
                  { label: "Accomplished Events", val: `${allEventsFlat.filter(ev => ev.status === "Accomplished").length}`, sub: `Out of ${totalEvents} total`, accent: "#0e9f6e" },
                  { label: "Total Programs", val: `${store.programs.length}`, sub: "Active programs", accent: "#7e3af2" },
                  { label: "Total Events", val: `${totalEvents}`, sub: `Across ${store.programs.length} program${store.programs.length === 1 ? "" : "s"}`, accent: "#d97706" },
                ].map((c, i) => (
                  <div key={i} style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 12, padding: "14px 16px" }}>
                    <div style={{ fontSize: 12, color: "#64748b", marginBottom: 4 }}>{c.label}</div>
                    <div style={{ fontSize: 26, fontWeight: 700, color: c.accent }}>{c.val}</div>
                    <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 2 }}>{c.sub}</div>
                  </div>
                ))}
              </div>

              {/* ── Target vs. Achievement by Program — Vertical Bar Chart ── */}
              <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 14, padding: "20px 24px", marginBottom: 28 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12, marginBottom: 20 }}>
                  <div>
                    <h2 style={{ fontSize: 16, fontWeight: 700, color: "#0f172a", margin: 0 }}>Events by Status per Program</h2>
                    <p style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>Count of Target (pending) vs. Accomplished events per program</p>
                  </div>
                </div>

                {chartData.length === 0 ? (
                  <div style={{ height: 240, display: "flex", alignItems: "center", justifyContent: "center", color: "#94a3b8", fontSize: 13 }}>No programs available to map.</div>
                ) : (
                  <div>
                    <div style={{ display: "flex", gap: 16, justifyContent: "flex-end", marginBottom: 16, fontSize: 12, fontWeight: 500 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <div style={{ width: 12, height: 12, background: "#cbd5e1", borderRadius: 3 }} />
                        <span style={{ color: "#475569" }}>Target</span>
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <div style={{ width: 12, height: 12, background: "#0f172a", borderRadius: 3 }} />
                        <span style={{ color: "#475569" }}>Accomplished</span>
                      </div>
                    </div>

                    <div style={{ display: "flex", height: 260, borderBottom: "2px solid #cbd5e1", paddingLeft: 10, position: "relative" }}>
                      {[0, 0.25, 0.5, 0.75, 1].map((ratio, index) => (
                        <div key={index} style={{ position: "absolute", bottom: `${ratio * 100}%`, left: 0, right: 0, borderTop: ratio === 0 ? "none" : "1px dashed #e2e8f0", height: 0 }}>
                          <span style={{ position: "absolute", left: 4, bottom: 2, fontSize: 10, color: "#94a3b8", fontWeight: 500 }}>{Math.round(ratio * yAxisMax)}</span>
                        </div>
                      ))}

                      <div style={{ display: "flex", width: "100%", justifyContent: "space-around", alignItems: "flex-end", zIndex: 1, paddingLeft: 36 }}>
                        {chartData.map((d, idx) => {
                          const targetHeight = (d.target / yAxisMax) * 100;
                          const actualHeight = (d.actual / yAxisMax) * 100;

                          return (
                            <div key={idx} style={{ display: "flex", flexDirection: "column", alignItems: "center", width: `${100 / chartData.length}%`, maxWidth: 110 }}>
                              <div style={{ display: "flex", alignItems: "flex-end", gap: 6, height: 220, width: "100%", justifyContent: "center" }}>
                                <div style={{ width: 22, height: `${Math.max(2, targetHeight)}%`, background: "#cbd5e1", borderRadius: "4px 4px 0 0", transition: "height 0.3s", position: "relative" }} title={`Target: ${d.target}`}>
                                  <span style={{ position: "absolute", top: -16, left: "50%", transform: "translateX(-50%)", fontSize: 10, fontWeight: 600, color: "#64748b" }}>{d.target}</span>
                                </div>
                                <div style={{ width: 22, height: `${Math.max(2, actualHeight)}%`, background: d.color, borderRadius: "4px 4px 0 0", transition: "height 0.3s", position: "relative" }} title={`Accomplished: ${d.actual}`}>
                                  <span style={{ position: "absolute", top: -16, left: "50%", transform: "translateX(-50%)", fontSize: 10, fontWeight: 700, color: d.color }}>{d.actual}</span>
                                </div>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>

                    <div style={{ display: "flex", justifyContent: "space-around", paddingLeft: 36, marginTop: 8 }}>
                      {chartData.map((d, idx) => (
                        <div key={idx} style={{ width: `${100 / chartData.length}%`, textAlign: "center", padding: "0 4px" }}>
                          <div style={{ fontSize: 11, fontWeight: 700, color: d.color, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }} title={d.code}>
                            {d.code}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              {/* Individual Program breakdown cards grid */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 16 }}>
                {store.programs.map(p => {
                  const ov = getOverall(p.id);
                  return (
                    <div key={p.id} style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 14, padding: "16px 18px" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
                        <div>
                          <div style={{ fontSize: 15, fontWeight: 700, color: p.color, display: "flex", alignItems: "center", gap: 6 }}><Dot color={p.color} /> {p.code}</div>
                          <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 2 }}>{p.name}</div>
                        </div>
                        <div style={{ fontSize: 22, fontWeight: 700, color: p.color }}>{ov !== null ? `${ov}%` : "—"}</div>
                      </div>
                      <ProgressBar value={ov} color={p.color} />
                      {(() => {
                        const cd = chartData.find(c => c.id === p.id);
                        return (
                          <div style={{ marginTop: 12, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                            <div style={{ background: "#f8fafc", borderRadius: 8, padding: "6px 8px", textAlign: "center" }}>
                              <div style={{ fontSize: 10, color: "#94a3b8" }}>Target</div>
                              <div style={{ fontSize: 14, fontWeight: 700, color: "#475569" }}>{cd ? cd.target : 0}</div>
                            </div>
                            <div style={{ background: "#f8fafc", borderRadius: 8, padding: "6px 8px", textAlign: "center" }}>
                              <div style={{ fontSize: 10, color: "#94a3b8" }}>Accomplished</div>
                              <div style={{ fontSize: 14, fontWeight: 700, color: p.color }}>{cd ? cd.actual : 0}</div>
                            </div>
                          </div>
                        );
                      })()}
                      <button onClick={() => { setActiveProgram(p.id); setView("programDetail"); }}
                        style={{ marginTop: 10, fontSize: 12, padding: "5px 12px", border: `1px solid ${p.color}`, borderRadius: 8, background: p.bg, color: p.color, cursor: "pointer", fontWeight: 500 }}>
                        <Eye size={12} style={{ verticalAlign: "-2px", marginRight: 4 }} /> View Events
                      </button>
                    </div>
                  );
                })}
              </div>

            {/* ── UPCOMING TARGET PROJECTS & EVENTS TABLE ── */}
              <div style={{ marginTop: 28, background: "#fff", border: "1px solid #e2e8f0", borderRadius: 14, padding: "20px 24px" }}>
                <div style={{ marginBottom: 16 }}>
                  <h2 style={{ fontSize: 16, fontWeight: 700, color: "#0f172a", margin: 0 }}>Upcoming Target Events</h2>
                  <p style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>Pending (Target) events, grouped by month · use the filter above to narrow the period</p>
                </div>

                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead>
                      <tr style={{ background: "#f8fafc", borderBottom: "1px solid #e2e8f0" }}>
                        <th style={{ padding: "12px 14px", textAlign: "left", fontWeight: 600, color: "#64748b" }}>Program</th>
                        <th style={{ padding: "12px 14px", textAlign: "left", fontWeight: 600, color: "#64748b" }}>Event / Task Name</th>
                        <th style={{ padding: "12px 14px", textAlign: "center", fontWeight: 600, color: "#64748b" }}>Date(s)</th>
                        <th style={{ padding: "12px 14px", textAlign: "left", fontWeight: 600, color: "#64748b" }}>Location</th>
                        <th style={{ padding: "12px 14px", textAlign: "left", fontWeight: 600, color: "#64748b" }}>Accompanying Personnel</th>
                        <th style={{ padding: "12px 14px", textAlign: "center", fontWeight: 600, color: "#64748b" }}></th>
                      </tr>
                    </thead>
                    <tbody>
                      {upcomingEventsByMonth.map(group => (
                        <Fragment key={group.key}>
                          <tr>
                            <td colSpan={6} style={{ padding: "10px 14px", background: "#f1f5f9", fontSize: 12, fontWeight: 700, color: "#334155" }}>
                              {group.label}
                            </td>
                          </tr>
                          {group.events.map((event, idx) => {
                            const associatedProg = store.programs.find(p => p.id === event.progId);
                            return (
                              <tr key={event.id || idx} style={{ borderBottom: "1px solid #f1f5f9", background: idx % 2 === 0 ? "#fff" : "#fafafa", transition: "background 0.15s" }}>
                                {/* Program Badge */}
                                <td style={{ padding: "14px 14px" }}>
                                  <span style={{
                                    background: associatedProg?.bg || "#f1f5f9",
                                    color: associatedProg?.color || "#475569",
                                    padding: "4px 10px",
                                    borderRadius: 6,
                                    fontSize: 12,
                                    fontWeight: 700
                                  }}>
                                    {associatedProg?.code || "DICT"}
                                  </span>
                                </td>

                                {/* Event name */}
                                <td style={{ padding: "14px 14px", fontWeight: 500, color: "#1e293b" }}>
                                  {event.name}
                                </td>

                                {/* Date string */}
                                <td style={{ padding: "14px 14px", textAlign: "center", color: "#475569", fontFamily: "monospace", fontWeight: 600 }}>
                                  {formatDateRange(event)}
                                </td>

                                {/* Location */}
                                <td style={{ padding: "14px 14px", color: "#64748b" }}>
                                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                    <MapPin size={13} style={{ color: "#94a3b8", flexShrink: 0 }} />
                                    <span>{event.location || "—"}</span>
                                  </div>
                                </td>

                                {/* Personnel names */}
                                <td style={{ padding: "14px 14px", color: "#64748b" }}>
                                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                    <div style={{ width: 24, height: 24, borderRadius: "50%", background: "#e2e8f0", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 700, color: "#475569" }}>
                                      👤
                                    </div>
                                    <span>{event.personnel || "Provincial Project Officer"}</span>
                                  </div>
                                </td>

                                {/* View / Expand Button */}
                                <td style={{ padding: "14px 14px", textAlign: "center" }}>
                                  <button onClick={() => setExpandedEventDetails(event)} title="View full details"
                                    style={{ border: "1px solid #e2e8f0", background: "#fff", color: "#475569", cursor: "pointer", padding: "5px 8px", borderRadius: 7, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
                                    <Eye size={14} />
                                  </button>
                                </td>
                              </tr>
                            );
                          })}
                        </Fragment>
                      ))}

                      {/* Fallback if no events have been encoded yet */}
                      {upcomingEventsByMonth.length === 0 && (
                        <tr>
                          <td colSpan={6} style={{ padding: "32px", textAlign: "center", color: "#94a3b8", fontStyle: "italic" }}>
                            No pending events for the selected period.
                                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ── SECTION B: ISOLATED TIMELINE SCHEDULE WORKSPACE ── */}
          {view === "schedule" && (
            <div>
              <div style={{ marginBottom: 24 }}>
                <h1 style={{ fontSize: 22, fontWeight: 700, margin: 0 }}>Program Milestones Schedule</h1>
                <p style={{ fontSize: 13, color: "#64748b", margin: "4px 0 0" }}>Map out specific objective deadlines across the operational calendar timeline</p>
              </div>

              {/* Target Monitoring Calendar Element Grid */}
              <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 14, padding: "20px 24px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12, marginBottom: 18 }}>
                  <div>
                    <h2 style={{ fontSize: 16, fontWeight: 700, color: "#0f172a", margin: 0 }}>📅 Milestone Operations Calendar</h2>
                    <p style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>Events shown here are pulled directly from each program's Events tab</p>
                  </div>

                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <button onClick={() => changeMonth("prev")} style={{ background: "#f1f5f9", border: "none", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontWeight: 600, fontSize: 12, color: "#475569" }}>◀</button>
                    <span style={{ fontSize: 14, fontWeight: 700, color: "#1e293b", minWidth: 120, textAlign: "center" }}>{MONTH_NAMES[currentCalendarMonth]} {currentCalendarYear}</span>
                    <button onClick={() => changeMonth("next")} style={{ background: "#f1f5f9", border: "none", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontWeight: 600, fontSize: 12, color: "#475569" }}>▶</button>
                  </div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "3fr 2fr", gap: 24, borderTop: "1px solid #f1f5f9", paddingTop: 16 }}>
                  <div>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", textAlign: "center", marginBottom: 8, fontWeight: 600, fontSize: 11, color: "#64748b" }}>
                      {["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"].map(day => <div key={day}>{day}</div>)}
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 6 }}>
                      {calendarDays.map((day, idx) => {
                        if (day === null) return <div key={`empty-${idx}`} style={{ height: 48 }} />;

                        const currentDayStr = `${currentCalendarYear}-${String(currentCalendarMonth + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
                        const dayEvents = eventsOnDate(currentDayStr);
                        const isDaySelected = selectedCalendarDate === currentDayStr;

                        return (
                          <div key={currentDayStr} onClick={() => setSelectedCalendarDate(currentDayStr)}
                            style={{ height: 52, padding: 4, background: isDaySelected ? "#eff6ff" : "#f8fafc", border: isDaySelected ? "2px solid #2563eb" : "1px solid #e2e8f0", borderRadius: 8, cursor: "pointer", display: "flex", flexDirection: "column", justifyContent: "space-between", position: "relative", transition: "all 0.15s" }}>
                            <span style={{ fontSize: 11, fontWeight: 700, color: isDaySelected ? "#1d4ed8" : "#475569" }}>{day}</span>

                            {dayEvents.length > 0 && (
                              <div style={{ display: "flex", gap: 2, overflow: "hidden", flexWrap: "wrap", maxWidth: "100%" }}>
                                {dayEvents.slice(0, 3).map((ev, mIdx) => {
                                  const pColor = store.programs.find(p => p.id === ev.progId)?.color || "#64748b";
                                  return (
                                    <div key={mIdx} style={{ width: 6, height: 6, borderRadius: "50%", background: pColor }} title={ev.name} />
                                  );
                                })}
                                {dayEvents.length > 3 && <span style={{ fontSize: 8, fontWeight: 700, color: "#94a3b8", lineHeight: 1 }}>+</span>}
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  <div style={{ background: "#f8fafc", borderRadius: 12, padding: 16, border: "1px solid #e2e8f0", display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
                    <div>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, borderBottom: "1px solid #e2e8f0", paddingBottom: 8 }}>
                        <span style={{ fontSize: 13, fontWeight: 700, color: "#0f172a" }}>🎯 Agenda: {selectedCalendarDate}</span>
                      </div>

                      <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 180, overflowY: "auto", paddingRight: 4 }}>
                        {eventsOnDate(selectedCalendarDate).length === 0 ? (
                          <div style={{ padding: "20px 0", textAlign: "center", color: "#94a3b8", fontSize: 12, fontStyle: "italic" }}>No events mapped to this day.</div>
                        ) : (
                          eventsOnDate(selectedCalendarDate).map((ev, mIdx) => {
                            const pItem = store.programs.find(p => p.id === ev.progId);
                            return (
                              <div key={ev.id || mIdx} style={{ background: "#fff", borderLeft: `4px solid ${pItem?.color || "#cbd5e1"}`, padding: "8px 12px", borderRadius: "0 8px 8px 0", boxShadow: "0 1px 2px rgba(0,0,0,0.05)", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
                                <div>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: pItem?.color }}>{pItem?.code}</div>
                                  <div style={{ fontSize: 12, fontWeight: 500, color: "#334155", marginTop: 2 }}>{ev.name}</div>
                                </div>
                                <button onClick={() => setExpandedEventDetails(ev)} title="View full details"
                                  style={{ border: "1px solid #e2e8f0", background: "#fff", color: "#475569", cursor: "pointer", padding: "5px 7px", borderRadius: 7, display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                                  <Eye size={13} />
                                </button>
                              </div>
                            );
                          })
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ── SECTION C: PROGRAMS LIST (table view) ── */}
          {view === "programs" && (
            <div>
              <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 20 }}>
                <div>
                  <h1 style={{ fontSize: 22, fontWeight: 700, margin: 0 }}>Programs</h1>
                  <p style={{ fontSize: 13, color: "#64748b", margin: "4px 0 0" }}>Manage DICT Abra programs and view their tracked events</p>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                  <PeriodFilter mode={periodMode} setMode={setPeriodMode} month={periodMonth} setMonth={setPeriodMonth} year={periodYear} setYear={setPeriodYear} />
                  <button onClick={openAddProgramModal}
                    style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: "#0f172a", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer", flexShrink: 0 }}>
                    <Plus size={16} /> Program
                  </button>
                </div>
              </div>

              <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 14, overflow: "hidden" }}>
                <div style={{ padding: "16px 20px", borderBottom: "1px solid #f1f5f9" }}>
                  <div style={{ position: "relative", maxWidth: 380 }}>
                    <Search size={15} style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: "#94a3b8" }} />
                    <input value={programSearch} onChange={e => setProgramSearch(e.target.value)} placeholder="Search programs by name or code..."
                      style={{ width: "100%", padding: "9px 12px 9px 34px", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 13, outline: "none", boxSizing: "border-box" }} />
                  </div>
                </div>

                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ background: "#f8fafc" }}>
                        <th style={th("left", "14%")}>Code</th>
                        <th style={th("left", "40%")}>Program Name</th>
                        <th style={th("center", "12%")}>Events</th>
                        <th style={th("center", "14%")}>Overall</th>
                        <th style={th("right", "20%")}>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredPrograms.length === 0 && (
                        <tr><td colSpan={5} style={{ textAlign: "center", padding: 32, color: "#94a3b8", fontSize: 13 }}>No programs match your search.</td></tr>
                      )}
                      {filteredPrograms.map((p, idx) => {
                        const ov = getOverall(p.id);
                        const eventCount = (store.events[p.id] || []).length;
                        return (
                          <tr key={p.id} style={{ borderTop: "1px solid #f1f5f9", background: idx % 2 === 0 ? "#fff" : "#fafafa" }}>
                            <td style={{ padding: "14px" }}>
                              <span style={{ display: "inline-flex", alignItems: "center", gap: 8, fontSize: 12, fontWeight: 700, color: p.color, background: p.bg, padding: "4px 10px", borderRadius: 6 }}>
                                <Dot color={p.color} /> {p.code}
                              </span>
                            </td>
                            <td style={{ padding: "14px", fontSize: 13, color: "#1e293b", fontWeight: 500 }}>{p.name}</td>
                            <td style={{ padding: "14px", textAlign: "center", fontSize: 13, color: "#64748b" }}>{eventCount}</td>
                            <td style={{ padding: "14px", textAlign: "center" }}>
                              <span style={{ fontSize: 13, fontWeight: 700, color: ov !== null ? (ov >= 90 ? "#0e9f6e" : ov >= 65 ? "#d97706" : "#e02424") : "#cbd5e1" }}>{ov !== null ? `${ov}%` : "—"}</span>
                            </td>
                            <td style={{ padding: "14px", textAlign: "right" }}>
                              <div style={{ display: "flex", gap: 6, justifyContent: "flex-end", alignItems: "center" }}>
                                <button onClick={() => { setActiveProgram(p.id); setView("programDetail"); }}
                                  style={{ display: "flex", alignItems: "center", gap: 5, padding: "6px 12px", border: `1px solid ${p.color}`, borderRadius: 7, background: p.bg, color: p.color, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                                  <Eye size={13} /> View
                                </button>
                                <button onClick={() => openEditProgramModal(p)} title="Edit program"
                                  style={{ border: "none", background: "none", color: "#64748b", cursor: "pointer", padding: 6, display: "flex" }}>
                                  <Pencil size={15} />
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ── SECTION C.2: EVENTS (Programs sub-tab) — all events, regardless of program/status ── */}
          {view === "events" && (
            <div>
              <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 20 }}>
                <div>
                  <h1 style={{ fontSize: 22, fontWeight: 700, margin: 0 }}>Events</h1>
                  <p style={{ fontSize: 13, color: "#64748b", margin: "4px 0 0" }}>All tracked events across every program</p>
                </div>
                <PeriodFilter mode={periodMode} setMode={setPeriodMode} month={periodMonth} setMonth={setPeriodMonth} year={periodYear} setYear={setPeriodYear} />
              </div>

              <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 14, overflow: "hidden" }}>
                <div style={{ padding: "16px 20px", borderBottom: "1px solid #f1f5f9" }}>
                  <div style={{ position: "relative", marginBottom: 12 }}>
                    <Search size={16} style={{ position: "absolute", left: 16, top: "50%", transform: "translateY(-50%)", color: "#94a3b8" }} />
                    <input value={eventsSearch} onChange={e => setEventsSearch(e.target.value)} placeholder="Search by event, employee, or location..."
                      style={{ width: "100%", padding: "11px 16px 11px 40px", border: "1px solid #e2e8f0", borderRadius: 999, fontSize: 13, outline: "none", boxSizing: "border-box", background: "#f8fafc" }} />
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", border: "1px solid #e2e8f0", borderRadius: 999, background: "#f8fafc" }}>
                      <Filter size={13} style={{ color: "#64748b", flexShrink: 0 }} />
                      <span style={{ fontSize: 12, color: "#64748b", fontWeight: 500 }}>From</span>
                      <input type="date" value={eventsDateFrom} onChange={e => setEventsDateFrom(e.target.value)}
                        style={{ border: "none", background: "transparent", fontSize: 12, outline: "none", color: "#334155", fontWeight: 500 }} />
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", border: "1px solid #e2e8f0", borderRadius: 999, background: "#f8fafc" }}>
                      <Filter size={13} style={{ color: "#64748b", flexShrink: 0 }} />
                      <span style={{ fontSize: 12, color: "#64748b", fontWeight: 500 }}>To</span>
                      <input type="date" value={eventsDateTo} onChange={e => setEventsDateTo(e.target.value)}
                        style={{ border: "none", background: "transparent", fontSize: 12, outline: "none", color: "#334155", fontWeight: 500 }} />
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", border: "1px solid #e2e8f0", borderRadius: 999, background: "#f8fafc" }}>
                      <Filter size={13} style={{ color: "#64748b", flexShrink: 0 }} />
                      <select value={eventsSort} onChange={e => setEventsSort(e.target.value)}
                        style={{ border: "none", background: "transparent", fontSize: 12, outline: "none", color: "#334155", fontWeight: 500, cursor: "pointer" }}>
                        <option value="date_asc">Date — Soonest first</option>
                        <option value="date_desc">Date — Latest first</option>
                        <option value="alpha">Alphabetical (A–Z)</option>
                      </select>
                    </div>
                  </div>
                </div>

                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead>
                      <tr style={{ background: "#f8fafc" }}>
                        <th style={th("left", "12%")}>Program</th>
                        <th style={th("left", "22%")}>Event</th>
                        <th style={th("center", "12%")}>Date</th>
                        <th style={th("left", "18%")}>Location</th>
                        <th style={th("left", "16%")}>Employee</th>
                        <th style={th("center", "10%")}>Status</th>
                        <th style={th("center", "5%")}></th>
                        <th style={th("center", "5%")}></th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredAllEvents.length === 0 && (
                        <tr><td colSpan={8} style={{ textAlign: "center", padding: 32, color: "#94a3b8", fontSize: 13 }}>No events match your filters.</td></tr>
                      )}
                      {filteredAllEvents.map((ev, idx) => {
                        const evProg = store.programs.find(p => p.id === ev.progId);
                        const isAccomplished = ev.status === "Accomplished";
                        return (
                          <tr key={ev.id} style={{ borderTop: "1px solid #f1f5f9", background: idx % 2 === 0 ? "#fff" : "#fafafa" }}>
                            <td style={{ padding: "10px 14px" }}>
                              <span style={{ background: evProg?.bg || "#f1f5f9", color: evProg?.color || "#475569", padding: "3px 8px", borderRadius: 6, fontSize: 11, fontWeight: 700 }}>
                                {evProg?.code || "—"}
                              </span>
                            </td>
                            <td style={{ padding: "10px 14px", fontWeight: 500, color: "#1e293b" }}>{ev.name}</td>
                            <td style={{ padding: "10px 8px", textAlign: "center", fontSize: 12, color: "#475569", fontFamily: "monospace" }}>{formatDateRange(ev)}</td>
                            <td style={{ padding: "10px 8px", fontSize: 12, color: "#475569" }}>{ev.location || "—"}</td>
                            <td style={{ padding: "10px 8px", fontSize: 12, color: "#475569" }}>{ev.personnel || "—"}</td>
                            <td style={{ padding: "10px 8px", textAlign: "center" }}>
                              <button onClick={() => toggleEventStatus(ev.progId, ev.id, ev.status)} title="Click to toggle status"
                                style={{
                                  fontSize: 11, fontWeight: 700, padding: "4px 8px", borderRadius: 10, border: "none", cursor: "pointer",
                                  background: isAccomplished ? "#def7ec" : "#eff6ff",
                                  color: isAccomplished ? "#03543f" : "#2563eb",
                                }}>
                                {ev.status || "Target"}
                              </button>
                            </td>
                            <td style={{ padding: "10px 8px", textAlign: "center" }}>
                              <button onClick={() => setExpandedEventDetails(ev)} title="View full details"
                                style={{ border: "1px solid #e2e8f0", background: "#fff", color: "#475569", cursor: "pointer", padding: "5px 7px", borderRadius: 7, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
                                <Eye size={13} />
                              </button>
                            </td>
                            <td style={{ padding: "10px 8px", textAlign: "center" }}>
                              <div style={{ display: "flex", gap: 4, justifyContent: "center" }}>
                                <button onClick={() => openEditEventModal(ev.progId, ev)} title="Edit event"
                                  style={{ border: "1px solid #e2e8f0", background: "#fff", color: "#475569", cursor: "pointer", padding: "5px 7px", borderRadius: 7, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
                                  <Pencil size={13} />
                                </button>
                                <button onClick={() => setDeleteEventConfirm({ progId: ev.progId, eventId: ev.id, name: ev.name })}
                                  style={{ border: "none", background: "none", color: "#ef4444", cursor: "pointer", fontSize: 14, opacity: .5, padding: 2 }} title="Remove event">✕</button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ── SECTION C.3: ADD EVENT (Events sub-tab) ── */}
          {view === "addEvent" && (
            <div>
              <div style={{ marginBottom: 24 }}>
                <h1 style={{ fontSize: 22, fontWeight: 700, margin: 0 }}>Add Event</h1>
                <p style={{ fontSize: 13, color: "#64748b", margin: "4px 0 0" }}>Create a new tracked event under a program.</p>
              </div>

              {store.programs.length === 0 ? (
                <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 14, padding: 40, textAlign: "center", color: "#94a3b8", fontSize: 13 }}>
                  No programs available yet. Create a program first under Programs.
                </div>
              ) : (
                <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 14, padding: "24px 28px", maxWidth: 520 }}>
                  <label style={labelStyle}>Program *</label>
                  <select value={activeProgram} onChange={e => setActiveProgram(e.target.value)} style={inputStyle}>
                    {store.programs.map(p => <option key={p.id} value={p.id}>{p.code} — {p.name}</option>)}
                  </select>

                  <label style={labelStyle}>Event Name *</label>
                  <input value={newEventName} onChange={e => setNewEventName(e.target.value)} placeholder="e.g., Beneficiaries trained"
                    style={inputStyle} />

                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                    <div>
                      <label style={labelStyle}>Start Date</label>
                      <input type="date" value={newEventStartDate} onChange={e => setNewEventStartDate(e.target.value)}
                        style={inputStyle} />
                    </div>
                    <div>
                      <label style={labelStyle}>End Date</label>
                      <input type="date" value={newEventEndDate} min={newEventStartDate || undefined} onChange={e => setNewEventEndDate(e.target.value)}
                        style={inputStyle} />
                    </div>
                  </div>

                  <label style={labelStyle}>Location</label>
                  <input value={newEventLocation} onChange={e => setNewEventLocation(e.target.value)} placeholder="e.g., Barangay Hall, Bangued, Abra"
                    style={inputStyle} />

                  <label style={labelStyle}>Assigned Employee/s</label>
                  <EmployeePicker employees={store.employees || []} value={newEventPersonnel} onChange={setNewEventPersonnel} />

                  <label style={labelStyle}>Description / Details</label>
                  <textarea value={newEventDescription} onChange={e => setNewEventDescription(e.target.value)} rows={3} placeholder="Brief description of this event..."
                    style={{ ...inputStyle, resize: "vertical" }} />

                  <label style={labelStyle}>Status</label>
                  <select value={newEventStatus} onChange={e => setNewEventStatus(e.target.value)}
                    style={{ ...inputStyle, marginBottom: 20 }}>
                    <option value="Target">Target</option>
                    <option value="Accomplished">Accomplished</option>
                  </select>

                  <div style={{ display: "flex", justifyContent: "flex-end" }}>
                    <button onClick={addEvent} disabled={!newEventName.trim()}
                      style={{ padding: "9px 22px", background: "#0f172a", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: newEventName.trim() ? "pointer" : "not-allowed", opacity: newEventName.trim() ? 1 : 0.5 }}>
                      <Plus size={14} style={{ verticalAlign: "-2px", marginRight: 4 }} /> Add Event
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* ── SECTION D: PROGRAM DETAIL — EVENTS DATA ENTRY VIEW ── */}
          {view === "programDetail" && (
            !prog ? (
              <div style={{ textAlign: "center", padding: 60, color: "#94a3b8" }}>
                <p style={{ marginBottom: 12 }}>No program selected.</p>
                <button onClick={() => setView("programs")} style={{ padding: "8px 18px", background: "#0f172a", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, cursor: "pointer" }}>Back to Programs</button>
              </div>
            ) : (
            <div>
              <button onClick={() => setView("programs")}
                style={{ display: "flex", alignItems: "center", gap: 6, border: "none", background: "none", color: "#64748b", fontSize: 12, fontWeight: 600, cursor: "pointer", padding: 0, marginBottom: 16 }}>
                <ArrowLeft size={14} /> Back to Programs
              </button>

              <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 20 }}>
                <div>
                  <h1 style={{ fontSize: 20, fontWeight: 700, margin: 0, color: prog.color, display: "flex", alignItems: "center", gap: 8 }}>
                    <Dot color={prog.color} size={10} /> {prog.code}
                  </h1>
                  <p style={{ fontSize: 12, color: "#64748b", margin: "3px 0 0" }}>{prog.name}</p>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                  <PeriodFilter mode={periodMode} setMode={setPeriodMode} month={periodMonth} setMonth={setPeriodMonth} year={periodYear} setYear={setPeriodYear} />
                  <button onClick={() => { setNoteText(store.notes?.[noteKey] || ""); setNotesOpen(true); }}
                    style={{ fontSize: 12, padding: "6px 12px", border: "1px solid #e2e8f0", borderRadius: 8, background: "#fff", color: "#475569", cursor: "pointer" }}>
                    📝 Notes / Remarks
                  </button>
                </div>
              </div>

              {/* Overall accomplishment summary widget */}
              {(() => {
                const ov = getOverall(activeProgram);
                return (
                  <div style={{ background: prog.bg, border: `1px solid ${prog.color}22`, borderRadius: 10, padding: "10px 16px", marginBottom: 16, display: "flex", alignItems: "center", gap: 16 }}>
                    <div>
                      <div style={{ fontSize: 11, color: "#64748b" }}>Overall Accomplishment Rate</div>
                      <div style={{ fontSize: 24, fontWeight: 700, color: prog.color }}>{ov !== null ? `${ov}%` : "No data yet"}</div>
                    </div>
                    {ov !== null && <ProgressBar value={ov} color={prog.color} />}
                    <div style={{ marginLeft: "auto" }}>{statusChip(ov)}</div>
                  </div>
                );
              })()}

              {/* Live Events Table Data Entry Panel */}
              <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 14, overflow: "hidden" }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ background: "#f8fafc" }}>
                      <th style={th("left", "26%")}>Event</th>
                      <th style={th("center", "14%")}>Date</th>
                      <th style={th("left", "18%")}>Location</th>
                      <th style={th("left", "18%")}>Employee</th>
                      <th style={th("center", "12%")}>Status</th>
                      <th style={th("center", "6%")}></th>
                      <th style={th("center", "6%")}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {events.length === 0 && (
                      <tr><td colSpan={7} style={{ textAlign: "center", padding: 32, color: "#94a3b8", fontSize: 13 }}>No events yet. Add one under Programs → Events.</td></tr>
                    )}
                    {events.map((ev, idx) => {
                      const isAccomplished = ev.status === "Accomplished";
                      return (
                        <tr key={ev.id} style={{ borderTop: "1px solid #f1f5f9", background: idx % 2 === 0 ? "#fff" : "#fafafa" }}>
                          <td style={{ padding: "10px 14px", fontSize: 13, fontWeight: 500, color: "#1e293b" }}>{ev.name}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", fontSize: 12, color: "#475569", fontFamily: "monospace" }}>{formatDateRange(ev)}</td>
                          <td style={{ padding: "10px 8px", fontSize: 12, color: "#475569" }}>{ev.location || "—"}</td>
                          <td style={{ padding: "10px 8px", fontSize: 12, color: "#475569" }}>{ev.personnel || "—"}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center" }}>
                            <button onClick={() => toggleEventStatus(activeProgram, ev.id, ev.status)} title="Click to toggle status"
                              style={{
                                fontSize: 11, fontWeight: 700, padding: "4px 8px", borderRadius: 10, border: "none", cursor: "pointer",
                                background: isAccomplished ? "#def7ec" : "#eff6ff",
                                color: isAccomplished ? "#03543f" : "#2563eb",
                              }}>
                              {ev.status || "Target"}
                            </button>
                          </td>
                          <td style={{ padding: "10px 8px", textAlign: "center" }}>
                            <button onClick={() => setExpandedEventDetails({ ...ev, progId: activeProgram })} title="View full details"
                              style={{ border: "1px solid #e2e8f0", background: "#fff", color: "#475569", cursor: "pointer", padding: "5px 7px", borderRadius: 7, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
                              <Eye size={13} />
                            </button>
                          </td>
                          <td style={{ padding: "10px 8px", textAlign: "center" }}>
                            <div style={{ display: "flex", gap: 4, justifyContent: "center" }}>
                              <button onClick={() => openEditEventModal(activeProgram, ev)} title="Edit event"
                                style={{ border: "1px solid #e2e8f0", background: "#fff", color: "#475569", cursor: "pointer", padding: "5px 7px", borderRadius: 7, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
                                <Pencil size={13} />
                              </button>
                              <button onClick={() => setDeleteEventConfirm({ progId: activeProgram, eventId: ev.id, name: ev.name })}
                                style={{ border: "none", background: "none", color: "#ef4444", cursor: "pointer", fontSize: 14, opacity: .5, padding: 2 }} title="Remove event">✕</button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
            )
          )}

          {/* ── SECTION E: FULL CONSOLIDATED REPORT VIEW ── */}
          {view === "report" && (
            <div>
              <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 20 }}>
                <div>
                  <h1 style={{ fontSize: 20, fontWeight: 700, margin: 0 }}>📋 Full Consolidated Report</h1>
                  <p style={{ fontSize: 12, color: "#64748b", margin: "3px 0 0" }}>All programs</p>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                  <PeriodFilter mode={periodMode} setMode={setPeriodMode} month={periodMonth} setMonth={setPeriodMonth} year={periodYear} setYear={setPeriodYear} />
                  <button onClick={exportToExcel}
                    style={{ padding: "7px 16px", background: "#1e40af", color: "#fff", border: "none", borderRadius: 8, fontWeight: 600, fontSize: 13, cursor: "pointer" }}>
                    📥 Export XLSX
                  </button>
                </div>
              </div>

              {store.programs.map(p => {
                const pevs = (store.events[p.id] || []).filter(ev => eventInPeriod(ev, periodMode, periodMonth, periodYear));
                const ov = getOverall(p.id);
                return (
                  <div key={p.id} style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 14, marginBottom: 20, overflow: "hidden" }}>
                    <div style={{ background: p.bg, padding: "12px 18px", borderBottom: "1px solid #e2e8f0", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                      <div>
                        <span style={{ fontSize: 15, fontWeight: 700, color: p.color }}>{p.code}</span>
                        <span style={{ fontSize: 12, color: "#64748b", marginLeft: 8 }}>{p.name}</span>
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                        <span style={{ fontSize: 12, color: "#64748b" }}>Overall:</span>
                        <span style={{ fontSize: 18, fontWeight: 700, color: ov !== null ? (ov >= 90 ? "#0e9f6e" : ov >= 65 ? "#d97706" : "#e02424") : "#94a3b8" }}>{ov !== null ? `${ov}%` : "—"}</span>
                        {statusChip(ov)}
                      </div>
                    </div>
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                        <thead>
                          <tr style={{ background: "#f8fafc" }}>
                            <th style={th("left", "30%")}>Event</th>
                            <th style={th("center", "14%")}>Date</th>
                            <th style={th("left", "26%")}>Location</th>
                            <th style={th("left", "18%")}>Employee</th>
                            <th style={th("center", "12%")}>Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          {pevs.length === 0 && (
                            <tr><td colSpan={5} style={{ textAlign: "center", padding: 20, color: "#94a3b8", fontSize: 12 }}>No events encoded.</td></tr>
                          )}
                          {pevs.map((ev, idx) => (
                            <tr key={ev.id} style={{ borderTop: "1px solid #f1f5f9", background: idx % 2 === 0 ? "#fff" : "#fafafa" }}>
                              <td style={{ padding: "8px 14px", fontWeight: 500, color: "#1e293b" }}>{ev.name}</td>
                              <td style={{ padding: "8px 6px", textAlign: "center", color: "#475569" }}>{formatDateRange(ev)}</td>
                              <td style={{ padding: "8px 6px", color: "#475569" }}>{ev.location || "—"}</td>
                              <td style={{ padding: "8px 6px", color: "#475569" }}>{ev.personnel || "—"}</td>
                              <td style={{ padding: "8px 6px", textAlign: "center" }}>
                                <span style={{
                                  fontSize: 10, fontWeight: 700, padding: "2px 7px", borderRadius: 8,
                                  background: ev.status === "Accomplished" ? "#def7ec" : "#eff6ff",
                                  color: ev.status === "Accomplished" ? "#03543f" : "#2563eb"
                                }}>{ev.status || "Target"}</span>
                              </td>
                            </tr>
                          ))}
                          {pevs.length > 0 && (
                            <tr style={{ background: p.bg, borderTop: "2px solid #e2e8f0", fontWeight: 700 }}>
                              <td colSpan={4} style={{ padding: "8px 14px", color: p.color }}>
                                Program Total ({pevs.filter(ev => ev.status === "Accomplished").length}/{pevs.length} Accomplished)
                              </td>
                              <td style={{ padding: "8px 8px", textAlign: "center" }}>{statusChip(ov)}</td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                    {store.notes?.[`${p.id}_notes`] && (
                      <div style={{ padding: "8px 16px", background: "#fffbeb", borderTop: "1px solid #fde68a", fontSize: 12, color: "#92400e" }}>
                        📝 <strong>Notes:</strong> {store.notes[`${p.id}_notes`]}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}

          {/* ── SECTION F: EMPLOYEES DIRECTORY ── */}
          {view === "employees" && (
            <div>
              <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 20 }}>
                <div>
                  <h1 style={{ fontSize: 22, fontWeight: 700, margin: 0 }}>Employees</h1>
                  <p style={{ fontSize: 13, color: "#64748b", margin: "4px 0 0" }}>Directory of personnel assigned across DICT Abra programs</p>
                </div>
                <button onClick={openAddEmployeeModal}
                  style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: "#0f172a", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer", flexShrink: 0 }}>
                  <Plus size={16} /> Employee
                </button>
              </div>

              {(store.employees || []).length === 0 ? (
                <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 14, padding: 40, textAlign: "center", color: "#94a3b8", fontSize: 13 }}>
                  No employees added yet.
                </div>
              ) : (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", gap: 16 }}>
                  {(store.employees || []).map((emp, idx) => {
                    const preset = COLOR_PRESETS[idx % COLOR_PRESETS.length];
                    const initials = (emp.name || "?").trim().split(/\s+/).slice(0, 2).map(w => w[0]).join("").toUpperCase();
                    return (
                      <div key={emp.id} style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 14, padding: "20px 16px", textAlign: "center", position: "relative" }}>
                        <div style={{ position: "absolute", top: 10, right: 10, display: "flex", gap: 2 }}>
                          <button onClick={() => openEditEmployeeModal(emp)} title="Edit employee"
                            style={{ border: "none", background: "none", color: "#94a3b8", cursor: "pointer", padding: 4, display: "flex" }}>
                            <Pencil size={13} />
                          </button>
                          <button onClick={() => setDeleteEmployeeConfirm(emp)} title="Remove employee"
                            style={{ border: "none", background: "none", color: "#94a3b8", cursor: "pointer", padding: 4, display: "flex" }}>
                            <Trash2 size={13} />
                          </button>
                        </div>
                        <div style={{ width: 64, height: 64, borderRadius: "50%", background: preset.bg, color: preset.color, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, fontWeight: 700, margin: "0 auto 12px" }}>
                          {initials}
                        </div>
                        <div style={{ fontSize: 14, fontWeight: 700, color: "#1e293b" }}>{emp.name}</div>
                        <div style={{ fontSize: 12, color: "#64748b", marginTop: 3 }}>{emp.role || "—"}</div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          )}

        </main>
      </div>

      {/* EDIT EVENT MODAL (opened via the pencil/edit button in a program's Events table) */}
      {editEventModalOpen && editEventForm && (() => {
        const editProg = store.programs.find(p => p.id === editEventForm.progId);
        const f = editEventForm;
        const set = (field, value) => setEditEventForm(prev => ({ ...prev, [field]: value }));
        return (
          <div onClick={() => setEditEventModalOpen(false)} style={overlayStyle}>
            <div onClick={e => e.stopPropagation()} style={{ background: "#fff", borderRadius: 16, padding: 28, width: "min(92vw, 440px)", maxHeight: "85vh", overflowY: "auto", boxShadow: "0 20px 60px rgba(0,0,0,.2)" }}>
              <h2 style={{ fontSize: 17, fontWeight: 700, margin: "0 0 4px", color: editProg?.color || "#0f172a" }}>Edit Event{editProg ? ` — ${editProg.code}` : ""}</h2>
              <p style={{ fontSize: 12, color: "#64748b", margin: "0 0 18px" }}>Update this event's details, location, and status.</p>

              <label style={labelStyle}>Event Name *</label>
              <input value={f.name} onChange={e => set("name", e.target.value)} style={inputStyle} />

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <div>
                  <label style={labelStyle}>Start Date</label>
                  <input type="date" value={f.startDate} onChange={e => set("startDate", e.target.value)} style={inputStyle} />
                </div>
                <div>
                  <label style={labelStyle}>End Date</label>
                  <input type="date" value={f.endDate} min={f.startDate || undefined} onChange={e => set("endDate", e.target.value)} style={inputStyle} />
                </div>
              </div>

              <label style={labelStyle}>Location</label>
              <input value={f.location} onChange={e => set("location", e.target.value)} placeholder="e.g., Barangay Hall, Bangued, Abra" style={inputStyle} />

              <label style={labelStyle}>Assigned Employee/s</label>
              <EmployeePicker employees={store.employees || []} value={f.personnel} onChange={val => set("personnel", val)} />

              <label style={labelStyle}>Description / Details</label>
              <textarea value={f.description} onChange={e => set("description", e.target.value)} rows={3} style={{ ...inputStyle, resize: "vertical" }} />

              <label style={labelStyle}>Status</label>
              <select value={f.status} onChange={e => set("status", e.target.value)}
                style={{ ...inputStyle, marginBottom: 20 }}>
                <option value="Target">Target</option>
                <option value="Accomplished">Accomplished</option>
              </select>

              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button onClick={() => setEditEventModalOpen(false)} style={cancelBtnStyle}>Cancel</button>
                <button onClick={saveEditEvent} disabled={!f.name.trim()}
                  style={{ padding: "8px 20px", background: editProg?.color || "#0f172a", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: f.name.trim() ? "pointer" : "not-allowed", opacity: f.name.trim() ? 1 : 0.5 }}>
                  Save Changes
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* EVENT DETAILS VIEW MODAL (view-only, opened via the eye/view button) */}
      {expandedEventDetails && (() => {
        const ev = expandedEventDetails;
        const evProg = store.programs.find(p => p.id === ev.progId);
        const isAccomplished = ev.status === "Accomplished";
        return (
          <div onClick={() => setExpandedEventDetails(null)} style={overlayStyle}>
            <div onClick={e => e.stopPropagation()} style={{ background: "#fff", borderRadius: 16, padding: 28, width: "min(92vw, 460px)", maxHeight: "80vh", overflowY: "auto", boxShadow: "0 20px 60px rgba(0,0,0,.2)" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                {evProg && (
                  <span style={{ background: evProg.bg, color: evProg.color, padding: "3px 9px", borderRadius: 6, fontSize: 11, fontWeight: 700 }}>
                    {evProg.code}
                  </span>
                )}
                <span style={{
                  background: isAccomplished ? "#def7ec" : "#eff6ff",
                  color: isAccomplished ? "#03543f" : "#2563eb",
                  fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 10
                }}>
                  {ev.status || "Target"}
                </span>
              </div>
              <h2 style={{ fontSize: 18, fontWeight: 700, margin: "8px 0 4px", color: "#0f172a" }}>{ev.name}</h2>
              {evProg && <p style={{ fontSize: 12, color: "#94a3b8", margin: "0 0 18px" }}>{evProg.name}</p>}

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 16 }}>
                <div style={{ background: "#f8fafc", borderRadius: 10, padding: "10px 12px" }}>
                  <div style={{ fontSize: 10, color: "#94a3b8", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 3 }}>Event Date(s)</div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: "#1e293b", fontFamily: "monospace" }}>{formatDateRange(ev)}</div>
                </div>
                <div style={{ background: "#f8fafc", borderRadius: 10, padding: "10px 12px" }}>
                  <div style={{ fontSize: 10, color: "#94a3b8", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 3 }}>Location</div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: "#1e293b" }}>{ev.location || "—"}</div>
                </div>
                <div style={{ background: "#f8fafc", borderRadius: 10, padding: "10px 12px" }}>
                  <div style={{ fontSize: 10, color: "#94a3b8", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 3 }}>Assigned Employee/s</div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: "#1e293b" }}>{ev.personnel || "Provincial Project Officer"}</div>
                </div>
              </div>

              <div style={{ marginBottom: 16 }}>
                <div style={{ fontSize: 10, color: "#94a3b8", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 5 }}>Description / Details</div>
                <div style={{ fontSize: 13, color: "#334155", lineHeight: 1.5, background: "#f8fafc", borderRadius: 10, padding: "10px 12px" }}>
                  {ev.description || "No description provided for this event."}
                </div>
              </div>

              <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 20 }}>
                <button onClick={() => setExpandedEventDetails(null)} style={cancelBtnStyle}>Close</button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* REMOVE EVENT CONFIRM MODAL */}
      {deleteEventConfirm && (
        <div onClick={() => setDeleteEventConfirm(null)} style={overlayStyle}>
          <div onClick={e => e.stopPropagation()} style={{ background: "#fff", borderRadius: 14, padding: 24, width: "min(92vw, 360px)" }}>
            <div style={{ fontSize: 32, textAlign: "center", marginBottom: 8 }}>⚠️</div>
            <h2 style={{ fontSize: 16, fontWeight: 700, textAlign: "center", margin: "0 0 8px" }}>Remove Event?</h2>
            <p style={{ fontSize: 13, color: "#475569", textAlign: "center", margin: "0 0 20px" }}>"{deleteEventConfirm.name}" and all its data will be permanently deleted.</p>
            <div style={{ display: "flex", gap: 8, justifyContent: "center" }}>
              <button onClick={() => setDeleteEventConfirm(null)} style={cancelBtnStyle}>Cancel</button>
              <button onClick={() => deleteEvent(deleteEventConfirm.progId, deleteEventConfirm.eventId)} style={{ padding: "8px 18px", background: "#ef4444", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Delete</button>
            </div>
          </div>
        </div>
      )}

      {/* REGIONAL PERFORMANCE NOTES MODAL */}
      {notesOpen && prog && (
        <div onClick={() => setNotesOpen(false)} style={overlayStyle}>
          <div onClick={e => e.stopPropagation()} style={{ background: "#fff", borderRadius: 16, padding: 28, width: "min(92vw, 440px)" }}>
            <h2 style={{ fontSize: 16, fontWeight: 700, margin: "0 0 4px" }}>📝 Notes / Remarks</h2>
            <p style={{ fontSize: 12, color: "#64748b", margin: "0 0 14px" }}>{prog.code}</p>
            <textarea value={noteText} onChange={e => setNoteText(e.target.value)} rows={5}
              placeholder="Enter remarks, explanations, or follow-up actions here..."
              style={{ width: "100%", padding: "10px 12px", border: "1px solid #d1d5db", borderRadius: 8, fontSize: 13, resize: "vertical", outline: "none", boxSizing: "border-box", marginBottom: 14 }} />
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setNotesOpen(false)} style={cancelBtnStyle}>Cancel</button>
              <button onClick={saveNote} style={{ padding: "8px 20px", background: "#1e40af", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Save Remarks</button>
            </div>
          </div>
        </div>
      )}

      {/* ADD / EDIT PROGRAM MODAL */}
      {programModalOpen && (
        <div onClick={() => setProgramModalOpen(false)} style={overlayStyle}>
          <div onClick={e => e.stopPropagation()} style={{ background: "#fff", borderRadius: 16, padding: 28, width: "min(92vw, 420px)", boxShadow: "0 20px 60px rgba(0,0,0,.2)" }}>
            <h2 style={{ fontSize: 17, fontWeight: 700, margin: "0 0 4px", color: "#0f172a" }}>
              {programModalMode === "add" ? "➕ Add Program" : "✏️ Edit Program"}
            </h2>
            <p style={{ fontSize: 12, color: "#64748b", margin: "0 0 18px" }}>
              {programModalMode === "add" ? "Create a new program to track events for." : "Update this program's details."}
            </p>

            <label style={labelStyle}>Acronym / Code *</label>
            <input value={programForm.code} onChange={e => setProgramForm(f => ({ ...f, code: e.target.value }))} placeholder="e.g., ILCDB"
              style={inputStyle} />

            <label style={labelStyle}>Full Program Name *</label>
            <input value={programForm.name} onChange={e => setProgramForm(f => ({ ...f, name: e.target.value }))} placeholder="e.g., ICT Literacy and Competency Development Bureau"
              style={inputStyle} />

            <label style={labelStyle}>Color</label>
            <div style={{ display: "flex", gap: 8, marginBottom: 22 }}>
              {COLOR_PRESETS.map((c, i) => (
                <button key={i} onClick={() => setProgramForm(f => ({ ...f, colorIdx: i }))}
                  style={{ width: 26, height: 26, borderRadius: "50%", background: c.color, border: programForm.colorIdx === i ? "2px solid #0f172a" : "2px solid #fff", boxShadow: "0 0 0 1px #e2e8f0", cursor: "pointer", padding: 0 }} />
              ))}
            </div>

            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setProgramModalOpen(false)} style={cancelBtnStyle}>Cancel</button>
              <button onClick={saveProgram} disabled={!programForm.code.trim() || !programForm.name.trim()}
                style={{ padding: "8px 20px", background: "#0f172a", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: (programForm.code.trim() && programForm.name.trim()) ? "pointer" : "not-allowed", opacity: (programForm.code.trim() && programForm.name.trim()) ? 1 : 0.5 }}>
                {programModalMode === "add" ? "Add Program" : "Save Changes"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ADD / EDIT EMPLOYEE MODAL */}
      {employeeModalOpen && (
        <div onClick={() => setEmployeeModalOpen(false)} style={overlayStyle}>
          <div onClick={e => e.stopPropagation()} style={{ background: "#fff", borderRadius: 16, padding: 28, width: "min(92vw, 400px)", boxShadow: "0 20px 60px rgba(0,0,0,.2)" }}>
            <h2 style={{ fontSize: 17, fontWeight: 700, margin: "0 0 4px", color: "#0f172a" }}>
              {employeeModalMode === "add" ? "➕ Add Employee" : "✏️ Edit Employee"}
            </h2>
            <p style={{ fontSize: 12, color: "#64748b", margin: "0 0 18px" }}>
              {employeeModalMode === "add" ? "Add a new employee to the directory." : "Update this employee's details."}
            </p>

            <label style={labelStyle}>Name *</label>
            <input value={employeeForm.name} onChange={e => setEmployeeForm(f => ({ ...f, name: e.target.value }))} placeholder="e.g., Jasmine A."
              style={inputStyle} />

            <label style={labelStyle}>Role</label>
            <input value={employeeForm.role} onChange={e => setEmployeeForm(f => ({ ...f, role: e.target.value }))} placeholder="e.g., Field Training Coordinator"
              style={{ ...inputStyle, marginBottom: 20 }} />

            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setEmployeeModalOpen(false)} style={cancelBtnStyle}>Cancel</button>
              <button onClick={saveEmployee} disabled={!employeeForm.name.trim()}
                style={{ padding: "8px 20px", background: "#0f172a", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: employeeForm.name.trim() ? "pointer" : "not-allowed", opacity: employeeForm.name.trim() ? 1 : 0.5 }}>
                {employeeModalMode === "add" ? "Add Employee" : "Save Changes"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* DELETE EMPLOYEE CONFIRM MODAL */}
      {deleteEmployeeConfirm && (
        <div onClick={() => setDeleteEmployeeConfirm(null)} style={overlayStyle}>
          <div onClick={e => e.stopPropagation()} style={{ background: "#fff", borderRadius: 14, padding: 24, width: "min(92vw, 360px)" }}>
            <div style={{ fontSize: 32, textAlign: "center", marginBottom: 8 }}>⚠️</div>
            <h2 style={{ fontSize: 16, fontWeight: 700, textAlign: "center", margin: "0 0 8px" }}>Remove Employee?</h2>
            <p style={{ fontSize: 13, color: "#475569", textAlign: "center", margin: "0 0 20px" }}>"{deleteEmployeeConfirm.name}" will be permanently removed from the directory.</p>
            <div style={{ display: "flex", gap: 8, justifyContent: "center" }}>
              <button onClick={() => setDeleteEmployeeConfirm(null)} style={cancelBtnStyle}>Cancel</button>
              <button onClick={() => deleteEmployee(deleteEmployeeConfirm.id)} style={{ padding: "8px 18px", background: "#ef4444", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Delete</button>
            </div>
          </div>
        </div>
      )}

      {/* ALERT ACTION CONFIRM NOTIFICATION TOAST */}
      {toast && (
        <div style={{ position: "fixed", bottom: 24, right: 24, background: toast.type === "success" ? "#0e9f6e" : toast.type === "info" ? "#1a56db" : "#e02424",
          color: "#fff", padding: "10px 20px", borderRadius: 10, fontSize: 13, fontWeight: 600, zIndex: 2000, boxShadow: "0 4px 20px rgba(0,0,0,.2)", animation: "fadeIn .2s" }}>
          {toast.msg}
        </div>
      )}

      <style>{`@keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}} input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none;}`}</style>
    </div>
  );
}

function th(align = "left", width) {
  return { padding: "9px 10px", textAlign: align, fontSize: 11, fontWeight: 600, color: "#64748b", borderBottom: "1px solid #e2e8f0", whiteSpace: "nowrap", width: width || "auto" };
}