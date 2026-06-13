import { useState, useEffect, useRef, useCallback } from "react";
import { useMsal } from "@azure/msal-react";
import {
  FileText,
  Download,
  Layers,
  AlertTriangle,
  Eye,
  ChevronLeft,
  ChevronRight,
  X,
  Activity,
  Brain,
  Cpu,
  Network,
  MessageSquare,
  CheckCircle2,
  Clock,
  TrendingUp,
  Shield,
  Users,
  BarChart3,
  Sparkles,
  Wrench,
  Scale,
  TrendingDown,
  Plus,
  Minus,
  RotateCcw,
  Calculator,
  SlidersHorizontal,
  ArrowRight,
  ArrowLeft,
  RefreshCcw,
  Gauge,
  PieChart,
  LineChart,
  AlertOctagon,
  Lightbulb,
  CheckCircle,
  Play,
  FastForward,
  Info,
  Loader2,
  FileSpreadsheet,
  ImageIcon,
  FileIcon,
  UploadCloud,
} from "lucide-react";
import logoArgunex from "./assets/logo_argunex.jpeg";

// ==========================================
// KONFIGURASI GLOBAL FRONTEND
// ==========================================
const WS_URL = "wss://YOUR_HF_SPACE_URL/ws"; // Ganti dengan URL WebSocket Backend Anda
const BACKEND_URL = "https://YOUR_HF_SPACE_URL"; // Ganti dengan URL Backend Anda
const MAX_FILE_SIZE = 2 * 1024 * 1024; // 2 MB

// ==========================================
// KOMPONEN MODAL ERROR API (VISUAL INDICATOR)
// ==========================================
function ApiErrorModal({ errorCode, onClose }) {
  if (!errorCode) return null;

  const errorConfig = {
    ERROR_401_UNAUTHORIZED: {
      title: "Autentikasi Gagal",
      message:
        "Sistem mengalami masalah autentikasi API Key. Harap hubungi administrator.",
      color: "text-red-600",
      bg: "bg-red-50",
      icon: Shield,
    },
    ERROR_402_NO_BALANCE: {
      title: "Kuota Habis",
      message:
        "Batas kuota simulasi operasional penuh (Saldo API Habis).",
      color: "text-amber-600",
      bg: "bg-amber-50",
      icon: AlertTriangle,
    },
    ERROR_429_LIMIT: {
      title: "Server Sibuk",
      message:
        "Server sedang sibuk memproses antrean data. Mohon tunggu 5 detik dan coba lagi.",
      color: "text-orange-600",
      bg: "bg-orange-50",
      icon: Clock,
    },
    ERROR_400_INVALID: {
      title: "Format Tidak Valid",
      message:
        "Permintaan gagal diproses karena format data atau parameter tidak valid.",
      color: "text-rose-600",
      bg: "bg-rose-50",
      icon: AlertOctagon,
    },
    ERROR_503_OVERLOAD: {
      title: "Infrastruktur Overload",
      message:
        "Infrastruktur DeepSeek global sedang mengalami kelebihan beban. Sistem otomatis beralih ke mode antrean.",
      color: "text-indigo-600",
      bg: "bg-indigo-50",
      icon: Network,
    },
  };

  const config = errorConfig[errorCode] || {
    title: "Koneksi Error",
    message: "Terjadi kesalahan koneksi ke server. Silakan coba lagi.",
    color: "text-slate-600",
    bg: "bg-slate-50",
    icon: AlertTriangle,
  };
  const Icon = config.icon;

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-[200] p-4">
      <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full p-8 text-center border border-slate-200">
        <div
          className={`w-16 h-16 rounded-full ${config.bg} flex items-center justify-center mx-auto mb-5`}
        >
          <Icon className={`w-8 h-8 ${config.color}`} />
        </div>
        <h3 className="text-xl font-bold text-slate-900 mb-3">
          {config.title}
        </h3>
        <p className="text-sm text-slate-600 mb-8 leading-relaxed">
          {config.message}
        </p>
        <button
          onClick={onClose}
          className="px-6 py-2.5 bg-[#4648d4] text-white rounded-xl font-semibold text-sm hover:bg-[#3638b0] transition-colors shadow-md"
        >
          Mengerti
        </button>
      </div>
    </div>
  );
}

// ==========================================
// KOMPONEN TOMBOL LOGIN MICROSOFT (REUSABLE)
// ==========================================
function MicrosoftLoginButton({ size = "small" }) {
  const { instance } = useMsal();

  const handleLogin = async () => {
    try {
      await instance.loginRedirect({
        scopes: ["user.read"],
        prompt: "select_account",
      });
    } catch (error) {
      console.error("Login error:", error);
    }
  };

  if (size === "large") {
    return (
      <button
        onClick={handleLogin}
        className="flex items-center gap-3 px-8 py-4 bg-[#0078d4] hover:bg-[#005a9e] text-white rounded-2xl text-lg font-bold transition-all shadow-xl hover:shadow-2xl active:scale-[0.98]"
      >
        <svg
          xmlns="http://www.w3.org/2000/svg"
          width="24"
          height="24"
          viewBox="0 0 21 21"
        >
          <rect x="1" y="1" width="9" height="9" fill="#f25022" />
          <rect x="1" y="11" width="9" height="9" fill="#00a4ef" />
          <rect x="11" y="1" width="9" height="9" fill="#7fba00" />
          <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
        </svg>
        Masuk dengan Microsoft
      </button>
    );
  }

  return (
    <button
      onClick={handleLogin}
      className="flex items-center gap-2 px-4 py-2 bg-[#0078d4] text-white rounded-lg text-xs font-bold hover:bg-[#005a9e] transition-colors shadow-sm"
    >
      <svg
        xmlns="http://www.w3.org/2000/svg"
        width="14"
        height="14"
        viewBox="0 0 21 21"
      >
        <rect x="1" y="1" width="9" height="9" fill="#f25022" />
        <rect x="1" y="11" width="9" height="9" fill="#00a4ef" />
        <rect x="11" y="1" width="9" height="9" fill="#7fba00" />
        <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
      </svg>
      Masuk dengan Microsoft
    </button>
  );
}

// ==========================================
// KOMPONEN DROPDOWN PROFIL USER (LOGOUT & GANTI AKUN)
// ==========================================
function UserProfileDropdown() {
  const { instance, accounts } = useMsal();
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef(null);

  const userName = accounts[0]?.name || "User";
  const firstName = userName.split(" ")[0];

  useEffect(() => {
    function handleClickOutside(e) {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target)) {
        setIsOpen(false);
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const getInitials = (name) => {
    if (!name) return "?";
    const names = name.split(" ");
    if (names.length === 1) return names[0].charAt(0).toUpperCase();
    return (
      names[0].charAt(0) + names[names.length - 1].charAt(0)
    ).toUpperCase();
  };

  const handleLogout = () => {
    setIsOpen(false);
    instance.logoutRedirect();
  };

  const handleSwitchAccount = () => {
    setIsOpen(false);
    instance.setActiveAccount(null);
    instance.loginRedirect({
      scopes: ["user.read"],
      prompt: "select_account",
    });
  };

  return (
    <div className="relative" ref={dropdownRef}>
      <div className="flex items-center gap-3">
        <span className="text-xs font-semibold text-slate-600 hidden md:block">
          Selamat Datang,{" "}
          <span className="text-[#4648d4] font-bold">{firstName}</span>
        </span>
        <div className="flex items-center gap-1.5 px-2.5 py-1 bg-emerald-50 text-emerald-700 rounded-full text-[10px] font-bold border border-emerald-200">
          <CheckCircle2 className="w-3 h-3" />
          Berhasil login akun microsoft
        </div>
        <button
          onClick={() => setIsOpen(!isOpen)}
          className="w-9 h-9 rounded-full bg-[#4648d4] text-white flex items-center justify-center font-bold text-xs shadow-md cursor-pointer border-2 border-white hover:bg-[#3638b0] transition-colors"
          title={userName}
        >
          {getInitials(userName)}
        </button>
      </div>

      {isOpen && (
        <div className="absolute right-0 mt-2 w-56 bg-white rounded-xl shadow-xl border border-slate-200 py-2 z-50 animate-fadeIn">
          <div className="px-4 py-3 border-b border-slate-100">
            <p className="text-sm font-bold text-slate-900">{userName}</p>
            <p className="text-xs text-slate-500 truncate">
              {accounts[0]?.username || ""}
            </p>
          </div>
          <button
            onClick={handleSwitchAccount}
            className="w-full text-left px-4 py-2.5 text-sm text-slate-700 hover:bg-slate-50 flex items-center gap-2 transition-colors"
          >
            <RefreshCcw className="w-4 h-4 text-slate-400" />
            Ganti Akun
          </button>
          <button
            onClick={handleLogout}
            className="w-full text-left px-4 py-2.5 text-sm text-red-600 hover:bg-red-50 flex items-center gap-2 transition-colors"
          >
            <ArrowLeft className="w-4 h-4 rotate-180 text-red-400" />
            Logout
          </button>
        </div>
      )}
    </div>
  );
}

// ==========================================
// HALAMAN LOGIN (AUTH GUARD)
// ==========================================
function AuthGuardScreen() {
  return (
    <div className="min-h-screen bg-[#f8f9fa] flex flex-col items-center justify-center px-6 hero-pattern">
      <div className="max-w-md w-full text-center">
        <img
          src={logoArgunex}
          alt="Argunex AI Logo"
          className="h-16 w-auto mx-auto mb-8 object-contain"
        />
        <h1 className="text-3xl font-extrabold text-[#191c1d] mb-4 tracking-tight">
          Selamat Datang di Argunex AI
        </h1>
        <p className="text-base text-[#464554] mb-10 leading-relaxed">
          Silakan login dengan akun Microsoft Anda untuk mengakses workspace dan
          fitur analisis multi-agent.
        </p>
        <MicrosoftLoginButton size="large" />
        <p className="mt-6 text-xs text-slate-400">
          Autentikasi diperlukan untuk melanjutkan.
        </p>
      </div>
    </div>
  );
}

// Helper functions
function getPhaseStyle(phase) {
  const styles = {
    idle: "bg-slate-100 text-slate-600 border-slate-200",
    roles: "bg-indigo-50 text-indigo-600 border-indigo-200",
    discussing: "bg-amber-50 text-amber-600 border-amber-200",
    moderating: "bg-rose-50 text-rose-600 border-rose-200",
    simulation: "bg-cyan-50 text-cyan-600 border-cyan-200",
    final: "bg-emerald-50 text-emerald-600 border-emerald-200",
  };
  return styles[phase] || styles.idle;
}

function getPhaseIcon(phase) {
  const icons = {
    idle: <Clock className="w-3.5 h-3.5" />,
    roles: <Users className="w-3.5 h-3.5" />,
    discussing: <MessageSquare className="w-3.5 h-3.5" />,
    moderating: <Shield className="w-3.5 h-3.5" />,
    simulation: <Calculator className="w-3.5 h-3.5" />,
    final: <CheckCircle2 className="w-3.5 h-3.5" />,
  };
  return icons[phase] || icons.idle;
}

// ==========================================
// FILE TYPE HELPERS
// ==========================================
function getFileTypeInfo(filename) {
  if (!filename)
    return { icon: FileIcon, color: "text-slate-500", label: "File" };
  const ext = filename.split(".").pop().toLowerCase();
  const map = {
    docx: { icon: FileText, color: "text-blue-500", label: "Word Document" },
    doc: { icon: FileText, color: "text-blue-500", label: "Word Document" },
    xlsx: {
      icon: FileSpreadsheet,
      color: "text-emerald-500",
      label: "Excel Spreadsheet",
    },
    xls: {
      icon: FileSpreadsheet,
      color: "text-emerald-500",
      label: "Excel Spreadsheet",
    },
    pdf: { icon: FileText, color: "text-red-500", label: "PDF Document" },
    png: { icon: ImageIcon, color: "text-purple-500", label: "Image (PNG)" },
    jpg: { icon: ImageIcon, color: "text-purple-500", label: "Image (JPG)" },
    jpeg: { icon: ImageIcon, color: "text-purple-500", label: "Image (JPEG)" },
  };
  return map[ext] || { icon: FileIcon, color: "text-slate-500", label: "File" };
}

// ==========================================
// POOL CONFIGURASI AGENT
// ==========================================
const AGENT_STYLE_POOL = [
  {
    color: "#3b82f6",
    bgColor: "#eff6ff",
    lucideIcon: FileText,
    subtitle: "Analyst",
  },
  {
    color: "#ef4444",
    bgColor: "#fef2f2",
    lucideIcon: Shield,
    subtitle: "Safety",
  },
  {
    color: "#8b5cf6",
    bgColor: "#f5f3ff",
    lucideIcon: Wrench,
    subtitle: "Engineer",
  },
  {
    color: "#10b981",
    bgColor: "#ecfdf5",
    lucideIcon: Scale,
    subtitle: "Legal",
  },
  {
    color: "#f59e0b",
    bgColor: "#fffbeb",
    lucideIcon: TrendingDown,
    subtitle: "Finance",
  },
  { color: "#ec4899", bgColor: "#fdf2f8", lucideIcon: Users, subtitle: "HRD" },
  {
    color: "#06b6d4",
    bgColor: "#ecfeff",
    lucideIcon: Network,
    subtitle: "Ops",
  },
  {
    color: "#f97316",
    bgColor: "#fff7ed",
    lucideIcon: Sparkles,
    subtitle: "Strategy",
  },
];

const STATIC_PARTICLES = Array.from({ length: 20 }, (_, i) => ({
  id: i,
  x: (i * 47) % 800,
  y: (i * 31) % 600,
  size: (i % 3) + 1.5,
  opacity: 0.12 + (i % 4) * 0.05,
}));

// ==========================================
// NETWORK GRAPH COMPONENT
// ==========================================
function NetworkGraph({
  zoomLevel,
  activeAgent,
  hoveredNode,
  setHoveredNode,
  problemText,
  agents,
}) {
  const containerRef = useRef(null);
  const [dimensions, setDimensions] = useState({ width: 800, height: 600 });

  useEffect(() => {
    const updateDim = () => {
      if (containerRef.current) {
        const rect = containerRef.current.getBoundingClientRect();
        setDimensions({ width: rect.width, height: rect.height });
      }
    };
    updateDim();
    window.addEventListener("resize", updateDim);
    return () => window.removeEventListener("resize", updateDim);
  }, []);

  const w = dimensions.width;
  const h = dimensions.height;
  const cx = w / 2;
  const cy = h / 2;
  const corePos = { x: cx, y: cy };

  const agentPositions = agents.map((agent, i) => {
    const angle =
      agents.length === 1
        ? -Math.PI / 2
        : (i / agents.length) * Math.PI * 2 - Math.PI / 2;
    const radius = Math.min(w, h) * 0.35;
    return {
      ...agent,
      x: cx + Math.cos(angle) * radius,
      y: cy + Math.sin(angle) * radius,
    };
  });

  return (
    <div
      ref={containerRef}
      className="absolute inset-0 overflow-hidden bg-slate-50"
    >
      <svg
        width={w}
        height={h}
        className="absolute inset-0"
        style={{
          transform: `scale(${zoomLevel})`,
          transformOrigin: "center center",
          transition: "transform 0.3s ease",
        }}
      >
        <defs>
          {agents.map((agent) => (
            <filter
              key={`glow-${agent.id}`}
              id={`glow-${agent.id}`}
              x="-50%"
              y="-50%"
              width="200%"
              height="200%"
            >
              <feGaussianBlur stdDeviation="6" result="blur" />
              <feMerge>
                <feMergeNode in="blur" />
                <feMergeNode in="SourceGraphic" />
              </feMerge>
            </filter>
          ))}
          <linearGradient id="coreGradient" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stopColor="#0891b2" stopOpacity="0.15" />
            <stop offset="100%" stopColor="#06b6d4" stopOpacity="0.3" />
          </linearGradient>
        </defs>
        {STATIC_PARTICLES.map((p) => (
          <circle
            key={p.id}
            cx={p.x}
            cy={p.y}
            r={p.size}
            fill="#94a3b8"
            opacity={p.opacity}
          />
        ))}
        {agentPositions.map((agent, i) => {
          const nextAgent =
            agents.length > 1
              ? agentPositions[(i + 1) % agentPositions.length]
              : null;
          return (
            <g key={`lines-${agent.id}`}>
              <line
                x1={agent.x}
                y1={agent.y}
                x2={corePos.x}
                y2={corePos.y}
                stroke={agent.color}
                strokeWidth="1.5"
                strokeDasharray="6 3"
                opacity="0.4"
                style={{ transition: "all 0.6s cubic-bezier(0.4, 0, 0.2, 1)" }}
              />
              <circle r="3" fill={agent.color} opacity="0.7">
                <animateMotion
                  dur="2.5s"
                  repeatCount="indefinite"
                  path={`M${agent.x},${agent.y} L${corePos.x},${corePos.y}`}
                />
              </circle>
              {nextAgent && (
                <line
                  x1={agent.x}
                  y1={agent.y}
                  x2={nextAgent.x}
                  y2={nextAgent.y}
                  stroke="#cbd5e1"
                  strokeWidth="1"
                  strokeDasharray="4 4"
                  opacity="0.3"
                  style={{
                    transition: "all 0.6s cubic-bezier(0.4, 0, 0.2, 1)",
                  }}
                />
              )}
            </g>
          );
        })}
        <g
          transform={`translate(${corePos.x}, ${corePos.y})`}
          style={{ transition: "all 0.3s ease" }}
        >
          <rect
            x="-95"
            y="-55"
            width="190"
            height="110"
            rx="16"
            fill="url(#coreGradient)"
            stroke="#06b6d4"
            strokeWidth="2"
          />
          <text
            x="0"
            y="-30"
            textAnchor="middle"
            fill="#0891b2"
            fontSize="10"
            fontWeight="800"
            fontFamily="monospace"
            letterSpacing="1.5"
          >
            CASE DATA ROOT
          </text>
          <text
            x="0"
            y="5"
            textAnchor="middle"
            fill="#0f172a"
            fontSize="12"
            fontWeight="700"
            className="select-none"
          >
            {problemText.length > 25
              ? `${problemText.substring(0, 25)}...`
              : problemText || "Awaiting Data..."}
          </text>
        </g>
        {agentPositions.map((agent) => {
          const isHovered = hoveredNode === agent.id;
          const isActive = activeAgent === agent.name;
          const nodeWidth = isHovered ? 165 : 150;
          const nodeHeight = isHovered ? 64 : 58;
          const IconComp = agent.lucideIcon;
          return (
            <g
              key={agent.id}
              style={{
                transform: `translate(${agent.x}px, ${agent.y}px)`,
                cursor: "pointer",
                transition: "transform 0.6s cubic-bezier(0.4, 0, 0.2, 1)",
              }}
              onMouseEnter={() => setHoveredNode(agent.id)}
              onMouseLeave={() => setHoveredNode(null)}
            >
              <rect
                x={-nodeWidth / 2}
                y={-nodeHeight / 2}
                width={nodeWidth}
                height={nodeHeight}
                rx="12"
                fill="#ffffff"
                stroke={agent.color}
                strokeWidth={isActive ? 3 : 1.5}
                filter={`url(#glow-${agent.id})`}
                style={{ transition: "all 0.2s ease" }}
              />
              <circle
                cx={-nodeWidth / 2 + 20}
                cy="0"
                r="14"
                fill={agent.bgColor}
                stroke={agent.color}
                strokeWidth="1"
              />
              <foreignObject
                x={-nodeWidth / 2 + 10}
                y={-10}
                width="20"
                height="20"
              >
                <div
                  style={{
                    color: agent.color,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    height: "100%",
                  }}
                >
                  <IconComp size={14} />
                </div>
              </foreignObject>
              <text
                x={-nodeWidth / 2 + 42}
                y="-2"
                fill="#0f172a"
                fontSize="10"
                fontWeight="800"
              >
                {agent.name}
              </text>
              <text
                x={-nodeWidth / 2 + 42}
                y="14"
                fill="#64748b"
                fontSize="9"
                fontFamily="monospace"
              >
                {agent.subtitle}
              </text>
            </g>
          );
        })}
      </svg>
    </div>
  );
}

// ==========================================
// SIMULATION ENGINE VIEW COMPONENT
// ==========================================
function SimulationView({
  simulationData,
  simulationResults,
  onRunSimulation,
  onConfirmSimulation,
  onSkipSimulation,
  onBackToDiscussion,
  wsConnected,
}) {
  const createInitialAdjustedValues = (variables = []) => {
    return variables.reduce((acc, variable) => {
      acc[variable.name] = variable.current_value;
      return acc;
    }, {});
  };

  const [adjustedValues, setAdjustedValues] = useState(() =>
    createInitialAdjustedValues(simulationData?.variables),
  );
  const [isRunning, setIsRunning] = useState(false);
  const [toast, setToast] = useState(null);

  useEffect(() => {
    if (!simulationData?.variables) return;
    const timer = setTimeout(() => {
      setAdjustedValues(createInitialAdjustedValues(simulationData.variables));
    }, 0);
    return () => clearTimeout(timer);
  }, [simulationData?.variables]);

  const handleSliderChange = (name, value) => {
    setAdjustedValues((prev) => ({ ...prev, [name]: parseFloat(value) }));
  };

  const handleReset = () => {
    if (!simulationData?.variables) return;
    const initial = {};
    simulationData.variables.forEach((v) => {
      initial[v.name] = v.current_value;
    });
    setAdjustedValues(initial);
  };

  const handleRun = () => {
    const changed = {};
    if (!simulationData?.variables) return;
    let isChanged = false;
    simulationData.variables.forEach((v) => {
      const current = adjustedValues[v.name] ?? v.current_value;
      if (Math.abs(current - v.current_value) > 0.0001) {
        changed[v.name] = current;
        isChanged = true;
      }
    });

    if (!isChanged) {
      setToast({
        type: "info",
        message: "No parameters adjusted. Running baseline simulation...",
      });
      simulationData.variables.forEach((v) => {
        changed[v.name] = adjustedValues[v.name] ?? v.current_value;
      });
    } else {
      setToast({
        type: "success",
        message: "Running simulation with adjusted parameters...",
      });
    }

    setTimeout(() => setToast(null), 3000);
    setIsRunning(true);
    onRunSimulation(changed);
  };

  const handleConfirm = () => {
    setIsRunning(false);
    onConfirmSimulation(simulationResults || {});
  };

  useEffect(() => {
    if (!simulationResults) return;
    const timeoutId = setTimeout(() => setIsRunning(false), 0);
    return () => clearTimeout(timeoutId);
  }, [simulationResults]);

  const formatNumber = (num) => {
    if (num === null || num === undefined) return "-";
    if (Math.abs(num) >= 1e9) return (num / 1e9).toFixed(2) + "B";
    if (Math.abs(num) >= 1e6) return (num / 1e6).toFixed(2) + "M";
    if (Math.abs(num) >= 1e3) return (num / 1e3).toFixed(1) + "K";
    if (Number.isInteger(num)) return num.toLocaleString("id-ID");
    return num.toLocaleString("id-ID", {
      minimumFractionDigits: 0,
      maximumFractionDigits: 2,
    });
  };

  const formatKpiName = (key) =>
    key
      .replace(/_/g, " ")
      .replace(/^-/, "")
      .trim()
      .replace(/\b\w/g, (l) => l.toUpperCase());

  const getRiskStyles = (level) => {
    switch (level) {
      case "high":
        return {
          bg: "bg-red-50 border-red-200",
          text: "text-red-700",
          icon: "text-red-500",
          badge: "bg-red-100 text-red-700",
        };
      case "elevated":
        return {
          bg: "bg-amber-50 border-amber-200",
          text: "text-amber-700",
          icon: "text-amber-500",
          badge: "bg-amber-100 text-amber-700",
        };
      default:
        return {
          bg: "bg-emerald-50 border-emerald-200",
          text: "text-emerald-700",
          icon: "text-emerald-500",
          badge: "bg-emerald-100 text-emerald-700",
        };
    }
  };

  const maxComparisonValue = (key) => {
    if (!simulationResults?.scenario_comparison?.[key]) return 1;
    const { original, adjusted } = simulationResults.scenario_comparison[key];
    return Math.max(Math.abs(original || 0), Math.abs(adjusted || 0), 1);
  };

  const hasVariables =
    simulationData?.variables && simulationData.variables.length > 0;

  return (
    <div className="h-[calc(100vh-72px)] flex flex-col bg-slate-50 overflow-hidden relative">
      {toast && (
        <div
          className={`fixed top-20 left-1/2 -translate-x-1/2 z-[100] px-5 py-3 rounded-xl shadow-lg flex items-center gap-3 text-sm font-medium border animate-slideDown ${toast.type === "info" ? "bg-indigo-50 text-indigo-700 border-indigo-200" : "bg-emerald-50 text-emerald-700 border-emerald-200"}`}
        >
          {toast.type === "info" ? (
            <Info className="w-4 h-4" />
          ) : (
            <CheckCircle2 className="w-4 h-4" />
          )}
          {toast.message}
        </div>
      )}

      <div className="bg-white border-b border-slate-200 px-6 py-4 flex items-center justify-between shrink-0">
        <div className="flex items-center gap-4">
          <button
            onClick={onBackToDiscussion}
            className="flex items-center gap-2 text-sm font-medium text-slate-500 hover:text-[#4648d4] transition-colors"
          >
            <ArrowLeft className="w-4 h-4" /> Kembali ke Analisis
          </button>
          <div className="h-6 w-px bg-slate-200" />
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-[#4648d4]/10 flex items-center justify-center">
              <Calculator className="w-4 h-4 text-[#4648d4]" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-slate-900">
                Simulation Engine
              </h2>
              <p className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold">
                Quantitative Scenario Modeling
              </p>
            </div>
          </div>
        </div>
        <div className="flex items-center gap-3">
          <div
            className={`flex items-center gap-2 px-3 py-1.5 rounded-full text-xs font-bold border ${wsConnected ? "bg-emerald-50 text-emerald-700 border-emerald-200" : "bg-red-50 text-red-700 border-red-200"}`}
          >
            <div
              className={`w-2 h-2 rounded-full ${wsConnected ? "bg-emerald-500 animate-pulse" : "bg-red-500"}`}
            />
            {wsConnected ? "Engine Online" : "Disconnected"}
          </div>
          <div className="px-3 py-1.5 bg-cyan-50 text-cyan-700 rounded-full text-xs font-bold border border-cyan-200">
            Phase 3: Simulation
          </div>
        </div>
      </div>

      <div className="flex-1 flex overflow-hidden">
        <div className="w-[420px] flex flex-col border-r border-slate-200 bg-white overflow-y-auto">
          <div className="p-6 border-b border-slate-100">
            <div className="flex items-center gap-2 mb-1">
              <Network className="w-4 h-4 text-[#4648d4]" />
              <span className="text-xs font-bold text-[#4648d4] uppercase tracking-wider">
                {simulationData?.domain || "Domain"}
              </span>
            </div>
            <h3 className="text-lg font-bold text-slate-900 leading-tight mb-2">
              {simulationData?.domain || "Simulation"}
            </h3>
            <p className="text-xs text-slate-500 leading-relaxed line-clamp-3">
              {simulationData?.original_problem || "No problem description"}
            </p>
            <div className="flex gap-3 mt-4">
              <div className="px-3 py-2 bg-slate-50 rounded-lg border border-slate-200">
                <p className="text-[10px] uppercase tracking-wider text-slate-500 font-bold">
                  Entitas
                </p>
                <p className="text-sm font-bold text-slate-900">
                  {simulationData?.entities?.count
                    ? `${simulationData.entities.count} ${simulationData.entities.name || ""}`
                    : "N/A"}
                </p>
              </div>
              <div className="px-3 py-2 bg-slate-50 rounded-lg border border-slate-200">
                <p className="text-[10px] uppercase tracking-wider text-slate-500 font-bold">
                  Variabel
                </p>
                <p className="text-sm font-bold text-slate-900">
                  {simulationData?.variables?.length || 0}
                </p>
              </div>
            </div>
          </div>

          <div className="p-6 flex-1">
            <div className="flex items-center justify-between mb-5">
              <h4 className="font-bold text-slate-900 text-sm flex items-center gap-2">
                <SlidersHorizontal className="w-4 h-4 text-[#4648d4]" />{" "}
                Parameter
              </h4>
              <button
                onClick={handleReset}
                className="text-xs text-slate-500 hover:text-[#4648d4] font-medium flex items-center gap-1 transition-colors"
              >
                <RefreshCcw className="w-3 h-3" /> Reset
              </button>
            </div>

            {hasVariables ? (
              <div className="space-y-5">
                {simulationData.variables.map((v, idx) => (
                  <div key={v.name} className="group">
                    <div className="flex items-center justify-between mb-2">
                      <div className="flex items-center gap-2">
                        <span className="w-5 h-5 rounded bg-[#4648d4]/10 text-[#4648d4] flex items-center justify-center text-[10px] font-bold">
                          {idx + 1}
                        </span>
                        <div>
                          <p className="text-sm font-semibold text-slate-900">
                            {formatKpiName(v.name)}
                          </p>
                          <p className="text-[10px] text-slate-500">
                            {v.description}
                          </p>
                        </div>
                      </div>
                      <div className="text-right">
                        <span className="text-sm font-bold text-[#4648d4] font-mono">
                          {formatNumber(
                            adjustedValues[v.name] ?? v.current_value,
                          )}
                        </span>
                        <span className="text-[10px] text-slate-500 ml-1">
                          {v.unit}
                        </span>
                      </div>
                    </div>
                    <div className="relative px-1">
                      <input
                        type="range"
                        min={v.min}
                        max={v.max}
                        step={v.step}
                        value={adjustedValues[v.name] ?? v.current_value}
                        onChange={(e) =>
                          handleSliderChange(v.name, e.target.value)
                        }
                        className="w-full h-1.5 bg-slate-200 rounded-full appearance-none cursor-pointer accent-[#4648d4]"
                      />
                      <div className="flex justify-between mt-1 text-[10px] text-slate-400 font-medium">
                        <span>{formatNumber(v.min)}</span>
                        <span className="text-slate-300">{v.scale}</span>
                        <span>{formatNumber(v.max)}</span>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <div className="text-center py-8">
                <AlertOctagon className="w-8 h-8 text-amber-400 mx-auto mb-3" />
                <p className="text-sm font-semibold text-slate-700 mb-1">
                  Tidak Ada Parameter
                </p>
                <p className="text-xs text-slate-500">
                  Variabel tidak dapat diekstrak. Anda tetap bisa lanjut ke rencana aksi.
                </p>
              </div>
            )}

            <div className="mt-6 space-y-3">
              <button
                onClick={handleRun}
                disabled={isRunning || !wsConnected}
                className="w-full py-3 bg-[#4648d4] hover:bg-[#3638b0] disabled:bg-slate-300 text-white rounded-xl font-semibold text-sm shadow-lg shadow-indigo-500/20 transition-all active:scale-[0.98] flex items-center justify-center gap-2"
              >
                {isRunning ? (
                  <RefreshCcw className="w-4 h-4 animate-spin" />
                ) : (
                  <Play className="w-4 h-4" />
                )}
                {isRunning ? "Menjalankan Simulasi..." : "Jalankan Simulasi"}
              </button>
              <button
                onClick={handleConfirm}
                disabled={!simulationResults || !wsConnected}
                className="w-full py-3 bg-emerald-500 hover:bg-emerald-600 disabled:bg-slate-300 text-white rounded-xl font-semibold text-sm shadow-lg shadow-emerald-500/20 transition-all active:scale-[0.98] flex items-center justify-center gap-2"
              >
                <CheckCircle className="w-4 h-4" /> Gunakan Hasil & Buat Blueprint
              </button>
              <button
                onClick={onSkipSimulation}
                disabled={!wsConnected}
                className="w-full py-3 bg-slate-100 hover:bg-slate-200 disabled:bg-slate-50 text-slate-700 rounded-xl font-semibold text-sm border border-slate-200 transition-all active:scale-[0.98] flex items-center justify-center gap-2"
              >
                <FastForward className="w-4 h-4" /> Lewati ke Blueprint
              </button>
            </div>
          </div>
        </div>

        <div className="flex-1 overflow-y-auto p-6 bg-slate-50">
          {!simulationResults ? (
            <div className="h-full flex flex-col items-center justify-center text-center">
              <div className="w-20 h-20 rounded-2xl bg-[#4648d4]/5 flex items-center justify-center mb-4">
                <Gauge className="w-8 h-8 text-[#4648d4]/40" />
              </div>
              <h3 className="text-lg font-semibold text-slate-900 mb-1">
                Hasil Simulasi
              </h3>
              <p className="text-sm text-slate-500 max-w-md">
                Sesuaikan parameter dan jalankan simulasi, atau lewati langsung ke Blueprint menggunakan data diskusi.
              </p>
            </div>
          ) : (
            <div className="max-w-5xl mx-auto space-y-6">
              {simulationResults.operational_impact &&
                (simulationResults.operational_impact.warnings?.length > 0 ||
                  simulationResults.operational_impact.recommendations?.length >
                    0) && (
                  <div
                    className={`rounded-2xl p-4 border flex items-start gap-3 ${getRiskStyles(simulationResults.operational_impact.risk_level).bg}`}
                  >
                    <AlertOctagon
                      className={`w-5 h-5 mt-0.5 shrink-0 ${getRiskStyles(simulationResults.operational_impact.risk_level).icon}`}
                    />
                    <div className="flex-1">
                      <div className="flex items-center gap-2 mb-2">
                        <span
                          className={`text-xs font-bold px-2 py-0.5 rounded-full ${getRiskStyles(simulationResults.operational_impact.risk_level).badge}`}
                        >
                          Risiko:{" "}
                          {(
                            simulationResults.operational_impact.risk_level ||
                            "normal"
                          ).toUpperCase()}
                        </span>
                      </div>
                      <div
                        className={`text-xs space-y-1 ${getRiskStyles(simulationResults.operational_impact.risk_level).text}`}
                      >
                        {simulationResults.operational_impact.warnings?.map(
                          (w, i) => (
                            <div key={i} className="flex items-start gap-1.5">
                              <span className="mt-1 w-1 h-1 rounded-full bg-current shrink-0" />
                              <span>{w}</span>
                            </div>
                          ),
                        )}
                      </div>
                      {simulationResults.operational_impact.recommendations
                        ?.length > 0 && (
                        <div
                          className={`text-xs mt-2 pt-2 border-t border-black/5 space-y-1 ${getRiskStyles(simulationResults.operational_impact.risk_level).text} opacity-80`}
                        >
                          {simulationResults.operational_impact.recommendations.map(
                            (r, i) => (
                              <div key={i} className="flex items-start gap-1.5">
                                <Lightbulb className="w-3 h-3 mt-0.5 shrink-0" />
                                <span>{r}</span>
                              </div>
                            ),
                          )}
                        </div>
                      )}
                    </div>
                  </div>
                )}

              <div>
                <h3 className="text-sm font-bold text-slate-900 mb-3 flex items-center gap-2">
                  <BarChart3 className="w-4 h-4 text-[#4648d4]" /> Indikator Kinerja Utama (KPI)
                </h3>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  {simulationResults.kpis &&
                  Object.keys(simulationResults.kpis).length > 0 ? (
                    Object.entries(simulationResults.kpis).map(
                      ([key, data]) => (
                        <div
                          key={key}
                          className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm hover:shadow-md transition-shadow"
                        >
                          <div className="flex items-center gap-2 mb-3">
                            <div className="w-8 h-8 rounded-lg bg-[#4648d4]/10 flex items-center justify-center">
                              <TrendingUp className="w-4 h-4 text-[#4648d4]" />
                            </div>
                            <span className="text-[10px] uppercase tracking-wider text-slate-500 font-bold">
                              {formatKpiName(key)}
                            </span>
                          </div>
                          <p className="text-2xl font-bold text-slate-900 mb-1">
                            {formatNumber(data.value)}
                          </p>
                          <p className="text-xs text-slate-500">{data.unit}</p>
                          <div className="mt-3 pt-3 border-t border-slate-100">
                            <p
                              className="text-[10px] text-slate-400 font-mono truncate"
                              title={data.formula}
                            >
                              {data.formula}
                            </p>
                          </div>
                        </div>
                      ),
                    )
                  ) : (
                    <div className="col-span-3 text-center py-8 text-slate-400 text-sm bg-white rounded-2xl border border-slate-200">
                      <BarChart3 className="w-6 h-6 mx-auto mb-2 opacity-40" />
                      <p>Tidak ada KPI untuk domain ini</p>
                    </div>
                  )}
                </div>
              </div>

              <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
                <h3 className="text-sm font-bold text-slate-900 mb-4 flex items-center gap-2">
                  <Calculator className="w-4 h-4 text-[#4648d4]" /> Perhitungan Dasar
                </h3>
                <div className="space-y-3">
                  {simulationResults.base_calculations &&
                  Object.keys(simulationResults.base_calculations).length >
                    0 ? (
                    Object.entries(simulationResults.base_calculations).slice(0, 10).map(
                      ([key, data]) => (
                        <div
                          key={key}
                          className="flex items-center gap-4 p-3 rounded-xl bg-slate-50 border border-slate-100"
                        >
                          <div className="w-10 h-10 rounded-lg bg-[#4648d4]/10 flex items-center justify-center shrink-0">
                            <span className="text-[#4648d4] font-bold text-sm">
                              =
                            </span>
                          </div>
                          <div className="flex-1 min-w-0">
                            <p className="text-xs font-semibold text-slate-900 uppercase tracking-wide">
                              {formatKpiName(key)}
                            </p>
                            <p className="text-xs text-slate-500 font-mono mt-0.5 truncate">
                              {data.formula} = {data.substitution}
                            </p>
                          </div>
                          <div className="text-right shrink-0">
                            <p className="text-lg font-bold text-[#4648d4] font-mono">
                              {formatNumber(data.result)}
                            </p>
                            <p className="text-[10px] text-slate-500">
                              {data.unit}
                            </p>
                          </div>
                        </div>
                      ),
                    )
                  ) : (
                    <p className="text-sm text-slate-400 italic">
                      Tidak ada perhitungan dasar
                    </p>
                  )}
                </div>
              </div>

              <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
                <h3 className="text-sm font-bold text-slate-900 mb-4 flex items-center gap-2">
                  <Scale className="w-4 h-4 text-[#4648d4]" /> Perbandingan Skenario
                </h3>
                <div className="space-y-5">
                  {simulationResults.scenario_comparison &&
                  Object.keys(simulationResults.scenario_comparison).length >
                    0 ? (
                    Object.entries(simulationResults.scenario_comparison).map(
                      ([key, data]) => {
                        if (
                          data.original === undefined &&
                          data.adjusted === undefined
                        )
                          return null;
                        const pct = data.change_percent || 0;
                        const isIncrease = data.impact_direction === "increase";
                        const color = isIncrease
                          ? pct > 50
                            ? "text-red-600"
                            : "text-emerald-600"
                          : "text-[#4648d4]";
                        const barColor = isIncrease
                          ? pct > 50
                            ? "bg-red-500"
                            : "bg-emerald-500"
                          : "bg-[#4648d4]";
                        const maxVal = maxComparisonValue(key);
                        const origPct =
                          (Math.abs(data.original || 0) / maxVal) * 100;
                        const adjPct =
                          (Math.abs(data.adjusted || 0) / maxVal) * 100;
                        return (
                          <div key={key} className="space-y-2">
                            <div className="flex items-center justify-between">
                              <span className="text-sm font-semibold text-slate-900">
                                {formatKpiName(key)}
                              </span>
                              <span
                                className={`text-xs font-bold flex items-center gap-1 ${color}`}
                              >
                                <TrendingUp
                                  className={`w-3 h-3 ${isIncrease ? "" : "rotate-180"}`}
                                />
                                {pct > 0 ? "+" : ""}
                                {pct}%
                              </span>
                            </div>
                            <div className="space-y-1.5">
                              <div className="flex items-center gap-3">
                                <span className="text-[10px] text-slate-500 w-14 text-right font-medium">
                                  Asli
                                </span>
                                <div className="flex-1 h-2 bg-slate-100 rounded-full overflow-hidden">
                                  <div
                                    className="h-full bg-slate-400 rounded-full transition-all duration-700"
                                    style={{
                                      width: `${Math.min(origPct, 100)}%`,
                                    }}
                                  />
                                </div>
                                <span className="text-xs font-mono text-slate-600 w-20 text-right">
                                  {formatNumber(data.original)}
                                </span>
                              </div>
                              <div className="flex items-center gap-3">
                                <span className="text-[10px] text-slate-500 w-14 text-right font-medium">
                                  Disesuaikan
                                </span>
                                <div className="flex-1 h-2 bg-slate-100 rounded-full overflow-hidden">
                                  <div
                                    className={`h-full ${barColor} rounded-full transition-all duration-700`}
                                    style={{
                                      width: `${Math.min(adjPct, 100)}%`,
                                    }}
                                  />
                                </div>
                                <span className="text-xs font-mono font-bold text-slate-800 w-20 text-right">
                                  {formatNumber(data.adjusted)}
                                </span>
                              </div>
                            </div>
                          </div>
                        );
                      },
                    )
                  ) : (
                    <div className="text-center py-6 text-slate-400 text-sm">
                      <p>Sesuaikan parameter untuk melihat perbandingan.</p>
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}


// ==========================================
// MAIN APP COMPONENT
// ==========================================
export default function App() {
  const { accounts } = useMsal();
  const isAuthenticated = accounts && accounts.length > 0;

  const [problemText, setProblemText] = useState("");
  const [clarificationText, setClarificationText] = useState("");
  
  const [ws, setWs] = useState(null);
  const [wsConnected, setWsConnected] = useState(false);
  
  const [phase, setPhase] = useState("idle");
  const [progress, setProgress] = useState(0);
  const [agents, setAgents] = useState([]);
  const [messages, setMessages] = useState([]);
  const [activeAgent, setActiveAgent] = useState("");
  
  const [moderatorQuestion, setModeratorQuestion] = useState("");
  const [simulationData, setSimulationData] = useState(null);
  const [simulationResults, setSimulationResults] = useState(null);
  const [finalBlueprint, setFinalBlueprint] = useState("");
  const [files, setFiles] = useState({ pdf: "", ppt: "" });
  
  const [apiError, setApiError] = useState(null);
  const [uploadingFile, setUploadingFile] = useState(false);
  const [uploadedFileData, setUploadedFileData] = useState(null);
  
  const [hoveredNode, setHoveredNode] = useState(null);
  const [zoomLevel] = useState(1);

  useEffect(() => {
    if (!isAuthenticated) return;
    
    const socket = new WebSocket(WS_URL);
    
    socket.onopen = () => {
      setWsConnected(true);
      console.log("WebSocket Connected");
    };
    
    socket.onclose = () => {
      setWsConnected(false);
      console.log("WebSocket Disconnected");
    };
    
    socket.onmessage = (event) => {
      const data = JSON.parse(event.data);
      
      if (data.step === "api_error") {
        setApiError(data.error_code);
        setPhase("idle");
        setProgress(0);
        return;
      }
      
      setProgress(data.progress || 0);
      
      switch(data.step) {
        case "roles":
          setPhase("roles");
          break;
        case "discussing":
          setPhase("discussing");
          if (data.agent && data.text) {
            setActiveAgent(data.agent);
            setMessages(prev => [...prev, { agent: data.agent, text: data.text }]);
            
            if (!agents.find(a => a.name === data.agent)) {
              const styleIdx = agents.length % AGENT_STYLE_POOL.length;
              const style = AGENT_STYLE_POOL[styleIdx];
              setAgents(prev => [...prev, {
                id: `agent-${prev.length}`,
                name: data.agent,
                color: style.color,
                bgColor: style.bgColor,
                lucideIcon: style.lucideIcon,
                subtitle: style.subtitle
              }]);
            }
          }
          break;
        case "moderating":
          setPhase("moderating");
          setActiveAgent("Moderator");
          break;
        case "ask_user":
          setPhase("moderating");
          setModeratorQuestion(data.text);
          break;
        case "simulation_ready":
          setPhase("simulation");
          setSimulationData(data.simulation_data);
          setSimulationResults(null);
          break;
        case "simulation_result":
          setSimulationResults(data.scenario_results);
          break;
        case "final":
          setPhase("final");
          setFinalBlueprint(data.content || "");
          setFiles(data.files || { pdf: "", ppt: "" });
          break;
      }
    };
    
    setWs(socket);
    
    return () => {
      socket.close();
    };
  }, [isAuthenticated]);

  const handleStartSimulation = () => {
    if (!ws || ws.readyState !== WebSocket.OPEN) return;
    setMessages([]);
    setAgents([]);
    setFinalBlueprint("");
    setSimulationData(null);
    setSimulationResults(null);
    setFiles({ pdf: "", ppt: "" });
    setPhase("roles");
    setProgress(5);
    
    let payload = problemText;
    if (uploadedFileData) {
      payload += `\n\nISI DOKUMEN (${uploadedFileData.filename}):\n${uploadedFileData.extracted_text}`;
    }
    
    ws.send(JSON.stringify({
      type: "start",
      problem: payload
    }));
  };

  const handleSendClarification = () => {
    if (!ws || ws.readyState !== WebSocket.OPEN || !clarificationText.trim()) return;
    ws.send(JSON.stringify({
      type: "answer",
      text: clarificationText.trim()
    }));
    setClarificationText("");
    setModeratorQuestion("");
  };

  const handleRunSimulation = (adjusted_params) => {
    if (!ws || ws.readyState !== WebSocket.OPEN) return;
    ws.send(JSON.stringify({
      type: "run_simulation",
      adjusted_params: adjusted_params
    }));
  };

  const handleConfirmSimulation = (sim_summary) => {
    if (!ws || ws.readyState !== WebSocket.OPEN) return;
    ws.send(JSON.stringify({
      type: "confirm_simulation",
      simulation_summary: sim_summary
    }));
  };

  const handleSkipSimulation = () => {
    if (!ws || ws.readyState !== WebSocket.OPEN) return;
    ws.send(JSON.stringify({ type: "skip_simulation" }));
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (file.size > MAX_FILE_SIZE) {
      setApiError("ERROR_400_INVALID"); // Using generic error UI for size limit
      return;
    }

    setUploadingFile(true);
    const formData = new FormData();
    formData.append("file", file);

    try {
      const res = await fetch(`${BACKEND_URL}/upload`, {
        method: "POST",
        body: formData,
      });
      
      if (!res.ok) {
        throw new Error("Upload failed");
      }
      
      const data = await res.json();
      setUploadedFileData(data);
    } catch (err) {
      console.error(err);
    } finally {
      setUploadingFile(false);
    }
  };

  if (!isAuthenticated) {
    return <AuthGuardScreen />;
  }

  // Render Final Blueprint & File Downloads
  if (phase === "final") {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col">
        <ApiErrorModal errorCode={apiError} onClose={() => setApiError(null)} />
        <header className="bg-white border-b border-slate-200 px-6 py-4 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <img src={logoArgunex} alt="Logo" className="h-8 w-auto" />
            <span className="font-bold text-slate-900">Argunex AI</span>
          </div>
          <UserProfileDropdown />
        </header>
        
        <div className="flex-1 max-w-5xl mx-auto w-full p-8">
          <div className="bg-white rounded-2xl shadow-lg border border-slate-200 p-8 mb-8">
            <div className="flex items-center gap-3 mb-6">
              <div className="w-10 h-10 rounded-xl bg-emerald-50 flex items-center justify-center">
                <CheckCircle2 className="w-6 h-6 text-emerald-600" />
              </div>
              <div>
                <h2 className="text-xl font-bold text-slate-900">Strategic Action Resolution Blueprint</h2>
                <p className="text-xs text-slate-500">Final output generated by Multi-Agent Committee</p>
              </div>
            </div>
            
            <div className="prose prose-sm max-w-none text-slate-700 whitespace-pre-wrap">
              {finalBlueprint}
            </div>
          </div>

          <div className="grid grid-cols-2 gap-6">
            <a
              href={files.pdf ? `${BACKEND_URL}${files.pdf}` : '#'}
              download="report.pdf"
              className={`flex items-center justify-center gap-3 p-5 rounded-xl border-2 transition-all ${files.pdf ? 'border-red-200 bg-red-50 hover:bg-red-100 text-red-700' : 'border-slate-200 bg-slate-50 text-slate-400 pointer-events-none'}`}
            >
              <FileText className="w-6 h-6" />
              <div className="text-left">
                <p className="font-bold text-sm">Download PDF</p>
                <p className="text-[10px] opacity-80">Laporan lengkap format dokumen</p>
              </div>
            </a>
            <a
              href={files.ppt ? `${BACKEND_URL}${files.ppt}` : '#'}
              download="report.pptx"
              className={`flex items-center justify-center gap-3 p-5 rounded-xl border-2 transition-all ${files.ppt ? 'border-orange-200 bg-orange-50 hover:bg-orange-100 text-orange-700' : 'border-slate-200 bg-slate-50 text-slate-400 pointer-events-none'}`}
            >
              <Layers className="w-6 h-6" />
              <div className="text-left">
                <p className="font-bold text-sm">Download PPTX</p>
                <p className="text-[10px] opacity-80">Presentasi strategi eksekutif</p>
              </div>
            </a>
          </div>

          <div className="mt-8 text-center">
            <button
              onClick={() => {
                setPhase("idle");
                setProblemText("");
                setMessages([]);
                setAgents([]);
                setFinalBlueprint("");
                setUploadedFileData(null);
                setProgress(0);
              }}
              className="px-6 py-3 bg-[#4648d4] text-white rounded-xl font-semibold text-sm hover:bg-[#3638b0] transition-colors shadow-md"
            >
              Mulai Sesi Baru
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (phase === "simulation") {
    return (
      <div className="min-h-screen bg-slate-50">
        <ApiErrorModal errorCode={apiError} onClose={() => setApiError(null)} />
        <SimulationView
          simulationData={simulationData}
          simulationResults={simulationResults}
          onRunSimulation={handleRunSimulation}
          onConfirmSimulation={handleConfirmSimulation}
          onSkipSimulation={handleSkipSimulation}
          onBackToDiscussion={() => setPhase("discussing")}
          wsConnected={wsConnected}
        />
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col">
      <ApiErrorModal errorCode={apiError} onClose={() => setApiError(null)} />
      
      <header className="bg-white border-b border-slate-200 px-6 py-4 flex items-center justify-between shrink-0 z-50">
        <div className="flex items-center gap-3">
          <img src={logoArgunex} alt="Logo" className="h-8 w-auto" />
          <span className="font-bold text-slate-900">Argunex AI</span>
        </div>
        <UserProfileDropdown />
      </header>

      <div className="flex-1 flex overflow-hidden">
        {/* Left Panel - Input & Chat */}
        <div className="w-[480px] flex flex-col border-r border-slate-200 bg-white">
          <div className="p-6 border-b border-slate-100">
            <h2 className="text-lg font-bold text-slate-900 mb-4 flex items-center gap-2">
              <Brain className="w-5 h-5 text-[#4648d4]" /> Input Masalah Operasional
            </h2>
            
            <div className="relative mb-4">
              <textarea
                rows={6}
                className="w-full p-3 border border-slate-200 rounded-xl text-sm text-slate-900 resize-none focus:ring-2 focus:ring-[#4648d4] focus:border-transparent transition-all"
                placeholder="Jelaskan masalah operasional yang ingin dianalisis..."
                value={problemText}
                onChange={(e) => setProblemText(e.target.value)}
                maxLength={1500}
                disabled={phase !== "idle"}
              />
              <div className="absolute bottom-3 right-3 text-[10px] text-slate-400 font-mono">
                {problemText.length}/1500
              </div>
            </div>

            <div className="flex items-center gap-3 mb-4">
              <label className={`flex-1 flex items-center justify-center gap-2 p-3 border-2 border-dashed ${uploadedFileData ? 'border-emerald-300 bg-emerald-50' : 'border-slate-200 hover:border-[#4648d4]'} rounded-xl cursor-pointer transition-colors ${phase !== "idle" ? 'pointer-events-none opacity-50' : ''}`}>
                <input 
                  type="file" 
                  className="hidden" 
                  accept=".docx,.doc,.xlsx,.xls,.pdf,.png,.jpg,.jpeg"
                  onChange={handleFileUpload}
                  disabled={phase !== "idle" || uploadingFile}
                />
                {uploadingFile ? (
                  <RefreshCcw className="w-4 h-4 animate-spin text-[#4648d4]" />
                ) : uploadedFileData ? (
                  <CheckCircle2 className="w-4 h-4 text-emerald-500" />
                ) : (
                  <UploadCloud className="w-4 h-4 text-slate-400" />
                )}
                <span className="text-xs font-semibold text-slate-600">
                  {uploadingFile ? "Mengunggah..." : uploadedFileData ? uploadedFileData.filename : "Unggah Dokumen (Maks 2MB)"}
                </span>
              </label>
            </div>

            <button
              onClick={handleStartSimulation}
              disabled={!problemText.trim() || !wsConnected || phase !== "idle"}
              className="w-full py-3 bg-[#4648d4] hover:bg-[#3638b0] disabled:bg-slate-300 text-white rounded-xl font-semibold text-sm shadow-lg transition-all active:scale-[0.98] flex items-center justify-center gap-2"
            >
              <Play className="w-4 h-4" /> Mulai Analisis Multi-Agent
            </button>
          </div>

          <div className="flex-1 overflow-y-auto p-6 space-y-4">
            {messages.map((msg, idx) => (
              <div key={idx} className="bg-slate-50 rounded-xl p-4 border border-slate-100">
                <p className="text-xs font-bold text-[#4648d4] mb-1">{msg.agent}</p>
                <p className="text-sm text-slate-700 leading-relaxed">{msg.text}</p>
              </div>
            ))}
            
            {moderatorQuestion && phase === "moderating" && (
              <div className="bg-amber-50 rounded-xl p-4 border border-amber-200">
                <div className="flex items-center gap-2 mb-2">
                  <AlertTriangle className="w-4 h-4 text-amber-500" />
                  <p className="text-xs font-bold text-amber-700 uppercase tracking-wider">Moderator Interruption</p>
                </div>
                <p className="text-sm text-amber-900 mb-4">{moderatorQuestion}</p>
                
                <div className="relative">
                  <textarea
                    rows={3}
                    className="w-full p-3 border border-amber-200 rounded-lg text-sm resize-none focus:ring-2 focus:ring-amber-400 focus:border-transparent bg-white"
                    placeholder="Jawab pertanyaan moderator..."
                    value={clarificationText}
                    onChange={(e) => setClarificationText(e.target.value)}
                    maxLength={500}
                  />
                  <div className="absolute bottom-2 right-2 text-[10px] text-slate-400 font-mono">
                    {clarificationText.length}/500
                  </div>
                </div>
                
                <button
                  onClick={handleSendClarification}
                  disabled={!clarificationText.trim()}
                  className="mt-2 w-full py-2 bg-amber-500 hover:bg-amber-600 disabled:bg-slate-300 text-white rounded-lg font-semibold text-xs transition-colors"
                >
                  Kirim Jawaban
                </button>
              </div>
            )}
          </div>
        </div>

        {/* Right Panel - Network Graph & Progress */}
        <div className="flex-1 flex flex-col relative">
          <div className="bg-white border-b border-slate-200 px-6 py-3 flex items-center justify-between z-10">
            <div className={`flex items-center gap-2 px-3 py-1.5 rounded-full text-xs font-bold border ${getPhaseStyle(phase)}`}>
              {getPhaseIcon(phase)}
              Phase: {phase.charAt(0).toUpperCase() + phase.slice(1)}
            </div>
            <div className="w-64 h-2 bg-slate-100 rounded-full overflow-hidden">
              <div 
                className="h-full bg-[#4648d4] rounded-full transition-all duration-500 ease-out"
                style={{ width: `${progress}%` }}
              />
            </div>
            <span className="text-xs font-bold text-slate-500">{progress}%</span>
          </div>

          <div className="flex-1 relative">
            <NetworkGraph
              zoomLevel={zoomLevel}
              activeAgent={activeAgent}
              hoveredNode={hoveredNode}
              setHoveredNode={setHoveredNode}
              problemText={problemText}
              agents={agents}
            />
          </div>
        </div>
      </div>
    </div>
  );
}
