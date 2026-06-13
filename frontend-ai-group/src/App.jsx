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
  Bot // Icon baru untuk modal
} from "lucide-react";
import logoArgunex from "./assets/logo_argunex.jpeg";

// ==========================================
// KONFIGURASI BACKEND URL (HUGGING FACE SPACE)
// ==========================================
const API_BASE = "https://muhammadfaiz2787-argunex-ai-backend.hf.space";
const WS_URL = "wss://muhammadfaiz2787-argunex-ai-backend.hf.space/ws";

// ==========================================
// KONSTANTA & UTILITAS
// ==========================================
const MAX_FILE_SIZE_FRONTEND = 2 * 1024 * 1024; // 2 MB
const MAX_INPUT_LENGTH = 1500;
const MAX_CLARIFICATION_LENGTH = 500;

const ERROR_MESSAGES = {
  ERROR_401_UNAUTHORIZED: "Sistem mengalami masalah autentikasi API Key. Harap hubungi administrator.",
  ERROR_402_NO_BALANCE: "Batas kuota simulasi operasional penuh (Saldo API Habis).",
  ERROR_429_LIMIT: "Server sedang sibuk memproses antrean data. Mohon tunggu 5 detik dan coba lagi.",
  ERROR_400_INVALID: "Format permintaan tidak valid. Silakan periksa input dan coba lagi.",
  ERROR_503_OVERLOAD: "Infrastruktur DeepSeek global sedang mengalami kelebihan beban. Sistem otomatis beralih ke mode antrean.",
};

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
        className="flex items-center justify-center w-fit mx-auto gap-3 px-8 py-4 bg-[#0078d4] hover:bg-[#005a9e] text-white rounded-2xl text-lg font-bold transition-all shadow-xl hover:shadow-2xl active:scale-[0.98]"
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

function getAgentCardStyle(status) {
  const styles = {
    idle: "border-slate-200 bg-slate-50",
    active: "border-amber-300 bg-amber-50 shadow-md shadow-amber-100",
    completed: "border-emerald-200 bg-emerald-50",
  };
  return styles[status] || styles.idle;
}

function getStatusDot(status) {
  const dots = {
    idle: "bg-slate-300",
    active: "bg-amber-500 animate-pulse",
    completed: "bg-emerald-500",
  };
  return dots[status] || dots.idle;
}

function getAgentStatusIcon(status) {
  switch (status) {
    case "active":
      return <Activity className="w-5 h-5 text-amber-500 animate-pulse" />;
    case "completed":
      return <CheckCircle2 className="w-5 h-5 text-emerald-500" />;
    default:
      return <Clock className="w-5 h-5 text-slate-400" />;
  }
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
  onNodeClick,
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
              onClick={() => onNodeClick && onNodeClick(agent.name)}
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
  const [isProblemExpanded, setIsProblemExpanded] = useState(false);

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
            <ArrowLeft className="w-4 h-4" /> Back to Analysis
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
            
            <div className="relative">
              <p className={`text-xs text-slate-500 leading-relaxed ${!isProblemExpanded ? 'line-clamp-3' : ''}`}>
                {simulationData?.original_problem || "No problem description"}
              </p>
              {(simulationData?.original_problem?.length > 150) && (
                <button
                  onClick={() => setIsProblemExpanded(!isProblemExpanded)}
                  className="text-[#4648d4] text-[10px] font-bold mt-1 hover:underline focus:outline-none"
                >
                  {isProblemExpanded ? "See Less" : "See More"}
                </button>
              )}
            </div>
            
            <div className="flex gap-3 mt-4">
              <div className="px-3 py-2 bg-slate-50 rounded-lg border border-slate-200">
                <p className="text-[10px] uppercase tracking-wider text-slate-500 font-bold">
                  Entities
                </p>
                <p className="text-sm font-bold text-slate-900">
                  {simulationData?.entities?.count
                    ? `${simulationData.entities.count} ${simulationData.entities.name || ""}`
                    : "N/A"}
                </p>
              </div>
              <div className="px-3 py-2 bg-slate-50 rounded-lg border border-slate-200">
                <p className="text-[10px] uppercase tracking-wider text-slate-500 font-bold">
                  Variables
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
                Parameters
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
                  No Adjustable Parameters
                </p>
                <p className="text-xs text-slate-500">
                  Variables could not be extracted. You can still proceed to the
                  action plan.
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
                {isRunning ? "Running Simulation..." : "Run Simulation"}
              </button>
              <button
                onClick={onSkipSimulation}
                disabled={!wsConnected}
                className="w-full py-3 bg-slate-100 hover:bg-slate-200 disabled:bg-slate-50 text-slate-700 rounded-xl font-semibold text-sm border border-slate-200 transition-all active:scale-[0.98] flex items-center justify-center gap-2"
              >
                <FastForward className="w-4 h-4" /> Skip to Action Plan
              </button>
            </div>
          </div>

          <div className="p-6 border-t border-slate-100 bg-slate-50/50">
            <h4 className="font-bold text-slate-900 text-sm flex items-center gap-2 mb-3">
              <Brain className="w-4 h-4 text-[#4648d4]" /> Expert Analysis
            </h4>
            <div className="text-xs text-slate-600 leading-relaxed max-h-40 overflow-y-auto space-y-2 pr-1">
              {simulationData?.expert_analysis ? (
                simulationData.expert_analysis
                  .split("\n")
                  .filter((l) => l.trim().length > 0)
                  .map((line, i) => {
                    const roleMatch = line.match(/^\[(.*?)\]:\s*(.*)$/);
                    if (roleMatch) {
                      const roleName = roleMatch[1];
                      const roleContent = roleMatch[2].trim();
                      return (
                        <div key={i} className="mb-2">
                          <div className="font-semibold text-[#4648d4] mt-2">
                            [{roleName}]:
                          </div>
                          {roleContent.length > 0 && (
                            <div className="pl-2 border-l-2 border-slate-200 text-xs text-slate-600 leading-relaxed">
                              {roleContent}
                            </div>
                          )}
                        </div>
                      );
                    }
                    return (
                      <div key={i} className="pl-2 border-l-2 border-slate-200 text-xs text-slate-600 leading-relaxed">
                        {line}
                      </div>
                    );
                  })
              ) : (
                <p className="text-slate-400 italic">
                  No expert analysis available
                </p>
              )}
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
                Simulation Results
              </h3>
              <p className="text-sm text-slate-500 max-w-md">
                Adjust parameters and run the simulation, or skip directly to
                the Action Plan using the discussion data.
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
                          Risk:{" "}
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
                  <BarChart3 className="w-4 h-4 text-[#4648d4]" /> Key
                  Performance Indicators
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
                      <p>No KPIs available for this domain</p>
                    </div>
                  )}
                </div>
              </div>

              <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
                <h3 className="text-sm font-bold text-slate-900 mb-4 flex items-center gap-2">
                  <Calculator className="w-4 h-4 text-[#4648d4]" /> Base
                  Calculations
                </h3>
                <div className="space-y-3">
                  {simulationResults.base_calculations &&
                  Object.keys(simulationResults.base_calculations).length >
                    0 ? (
                    Object.entries(simulationResults.base_calculations).map(
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
                            <p className="text-xs text-slate-500 font-mono mt-0.5">
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
                      No base calculations available
                    </p>
                  )}
                </div>
              </div>

              <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
                <h3 className="text-sm font-bold text-slate-900 mb-4 flex items-center gap-2">
                  <Scale className="w-4 h-4 text-[#4648d4]" /> Scenario
                  Comparison
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
                                  Original
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
                                  Adjusted
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
                      <p>
                        No comparison data. Adjust parameters to see
                        comparisons.
                      </p>
                    </div>
                  )}
                </div>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
                  <h4 className="font-bold text-slate-900 text-sm mb-4 flex items-center gap-2">
                    <LineChart className="w-4 h-4 text-[#4648d4]" /> Parameter
                    Impact (%)
                  </h4>
                  <div className="space-y-3">
                    {simulationResults.scenario_comparison &&
                      Object.entries(simulationResults.scenario_comparison)
                        .filter(([, data]) => data.change_percent !== undefined)
                        .map(([key, data]) => {
                          const pct = data.change_percent || 0;
                          const isPos = pct >= 0;
                          const color = isPos
                            ? pct > 50
                              ? "bg-red-500"
                              : "bg-emerald-500"
                            : "bg-[#4648d4]";
                          const barWidth = Math.min(Math.abs(pct), 100);
                          return (
                            <div key={key} className="flex items-center gap-3">
                              <span className="text-xs text-slate-600 w-32 truncate font-medium">
                                {formatKpiName(key)}
                              </span>
                              <div className="flex-1 h-6 bg-slate-50 rounded-lg overflow-hidden relative">
                                <div
                                  className={`absolute top-0 h-full ${color} rounded-lg transition-all duration-700 flex items-center justify-end px-2`}
                                  style={{
                                    width: `${barWidth}%`,
                                    left: isPos ? "0" : `${100 - barWidth}%`,
                                  }}
                                >
                                  {barWidth > 15 && (
                                    <span className="text-[10px] font-bold text-white">
                                      {isPos ? "+" : ""}
                                      {pct}%
                                    </span>
                                  )}
                                </div>
                                {barWidth <= 15 && (
                                  <span className="absolute inset-0 flex items-center pl-2 text-[10px] font-bold text-slate-600">
                                    {isPos ? "+" : ""}
                                    {pct}%
                                  </span>
                                )}
                              </div>
                            </div>
                          );
                        })}
                  </div>
                </div>
                <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
                  <h4 className="font-bold text-slate-900 text-sm mb-4 flex items-center gap-2">
                    <PieChart className="w-4 h-4 text-[#4648d4]" /> Value
                    Comparison
                  </h4>
                  <div className="space-y-3">
                    {simulationResults.scenario_comparison &&
                      Object.entries(simulationResults.scenario_comparison)
                        .filter(
                          ([, data]) =>
                            data.original !== undefined &&
                            data.adjusted !== undefined,
                        )
                        .slice(0, 6)
                        .map(([key, data]) => {
                          const max = Math.max(
                            Math.abs(data.original),
                            Math.abs(data.adjusted),
                            1,
                          );
                          return (
                            <div key={key}>
                              <div className="flex items-center justify-between mb-1">
                                <span className="text-xs text-slate-600 font-medium">
                                  {formatKpiName(key)}
                                </span>
                                <span className="text-[10px] text-slate-400">
                                  {formatNumber(data.adjusted)} vs{" "}
                                  {formatNumber(data.original)}
                                </span>
                              </div>
                              <div className="flex items-center gap-1 h-4">
                                <div
                                  className="h-full bg-slate-300 rounded-l-md transition-all duration-700"
                                  style={{
                                    width: `${(Math.abs(data.original) / max) * 50}%`,
                                  }}
                                />
                                <div
                                  className="h-full bg-[#4648d4] rounded-r-md transition-all duration-700"
                                  style={{
                                    width: `${(Math.abs(data.adjusted) / max) * 50}%`,
                                  }}
                                />
                              </div>
                            </div>
                          );
                        })}
                  </div>
                </div>
              </div>

              <div className="bg-gradient-to-br from-[#4648d4]/5 to-white rounded-2xl p-6 border border-[#4648d4]/20 shadow-sm">
                <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
                  <div>
                    <h4 className="font-bold text-slate-900 text-sm">
                      Proceed to Action Plan?
                    </h4>
                    <p className="text-xs text-slate-600 mt-1">
                      Simulation results will be used as the basis for strategic
                      recommendations and SOP.
                    </p>
                  </div>
                  <button
                    onClick={() => onConfirmSimulation(simulationResults)}
                    className="px-6 py-3 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-semibold text-sm shadow-lg shadow-emerald-500/20 transition-all active:scale-[0.98] flex items-center gap-2 whitespace-nowrap"
                  >
                    <CheckCircle className="w-4 h-4" /> Confirm & Proceed{" "}
                    <ArrowRight className="w-4 h-4" />
                  </button>
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
  const [view, setView] = useState("home");
  const [input, setInput] = useState("");
  const [progress, setProgress] = useState(0);
  const [messages, setMessages] = useState([]);
  const [isWaitingUser, setIsWaitingUser] = useState(false);
  const [modQuestion, setModQuestion] = useState("");
  const [selectedFile, setSelectedFile] = useState(null);
  const [extractedDocText, setExtractedDocText] = useState("");
  const [isProcessingFile, setIsProcessingFile] = useState(false);
  const [fileUploadError, setFileUploadError] = useState("");
  const [isDraggingOver, setIsDraggingOver] = useState(false);
  const [result, setResult] = useState({ content: "", files: {} });
  const [currentLog, setCurrentLog] = useState(
    "System initialized. Awaiting user input...",
  );
  const [pdfPreviewUrl, setPdfPreviewUrl] = useState(null);
  const [pptSlides, setPptSlides] = useState([]);
  const [isPptPreviewOpen, setIsPptPreviewOpen] = useState(false);
  const [currentSlideIndex, setCurrentSlideIndex] = useState(0);
  const [activeAgent, setActiveAgent] = useState(null);
  const [discussionPhase, setDiscussionPhase] = useState("idle");
  const [problemText, setProblemText] = useState("");
  const [simulationData, setSimulationData] = useState(null);
  const [simulationResults, setSimulationResults] = useState(null);
  const [wsConnected, setWsConnected] = useState(false);
  const [zoomLevel, setZoomLevel] = useState(1);
  const [hoveredNode, setHoveredNode] = useState(null);
  const [apiError, setApiError] = useState(null);
  const [selectedAgentText, setSelectedAgentText] = useState(null); // State untuk modal discussion

  const ws = useRef(null);
  const reconnectTimer = useRef(null);

  const forceDownload = async (url, filename) => {
    try {
      setCurrentLog(`⬇️ Downloading ${filename}...`);
      const response = await fetch(url, { method: "GET" });
      if (!response.ok) throw new Error(`Server returned ${response.status}`);
      const blob = await response.blob();
      const blobUrl = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = blobUrl;
      a.download = filename;
      a.style.display = "none";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(blobUrl);
      setCurrentLog(`✅ ${filename} downloaded successfully.`);
    } catch (err) {
      console.error("Download failed:", err);
      setCurrentLog(`❌ Download failed: ${err.message}. Try using the Preview button.`);
    }
  };

  const uniqueAgentNames = Array.from(
    new Set(messages.map((m) => m.agent).filter(Boolean)),
  );
  const visibleAgents = uniqueAgentNames.map((name, idx) => {
    const styleObj = AGENT_STYLE_POOL[idx % AGENT_STYLE_POOL.length];
    return { ...styleObj, name: name, id: `dyn_agent_${idx}` };
  });

  const handleWsMessage = useCallback((event) => {
    try {
      const data = JSON.parse(event.data);
      if (data.progress) setProgress(data.progress);

      if (data.step === "error") {
        const errCode = data.error_code || "ERROR_503_OVERLOAD";
        if (errCode === "ERROR_401_UNAUTHORIZED" || errCode === "ERROR_402_NO_BALANCE") {
          setApiError(errCode);
          setDiscussionPhase("idle");
          setProgress(0);
          setView("workspace");
        } else {
          setCurrentLog(`Transient error: ${errCode}. Retrying...`);
        }
        return;
      }

      if (data.step === "roles") {
        setCurrentLog("System assigning optimized multi-agent roles...");
        setDiscussionPhase("roles");
      } else if (data.step === "discussing") {
        setMessages((prev) => [...prev, { agent: data.agent, text: data.text }]);
        setActiveAgent(data.agent);
        setCurrentLog(`${data.agent}: ${data.text.substring(0, 80)}...`);
        setDiscussionPhase("discussing");
      } else if (data.step === "ask_user") {
        setIsWaitingUser(true);
        setModQuestion(data.text);
        setCurrentLog("Moderator interaction requested. Discussion paused.");
        setDiscussionPhase("moderating");
        setActiveAgent("user");
      } else if (data.step === "simulation_ready") {
        setSimulationData(data.simulation_data);
        setSimulationResults(null);
        setView("simulation");
        setCurrentLog(
          "Simulation engine initialized. Ready for scenario modeling.",
        );
        setDiscussionPhase("simulation");
        setActiveAgent(null);
      } else if (data.step === "simulation_result") {
        setSimulationResults(data.scenario_results);
        setCurrentLog("Simulation completed. Scenario analysis ready.");
      } else if (data.step === "final") {
        setResult({ content: data.content, files: data.files });
        if (data.slides_preview) setPptSlides(data.slides_preview);
        setView("summary");
        setCurrentLog(
          "Analysis completed. Final quantitative reports generated.",
        );
        setDiscussionPhase("final");
        setActiveAgent(null);
      }
    } catch (parseErr) {
      console.error("[WS PARSE ERROR]", parseErr);
      setCurrentLog("Received malformed message from server. Ignoring...");
    }
  }, []);

  const connectWebSocket = useCallback(() => {
    if (ws.current?.readyState === WebSocket.OPEN || ws.current?.readyState === WebSocket.CONNECTING) {
      return;
    }

    try {
      const wsInstance = new WebSocket(WS_URL);
      wsInstance.onmessage = handleWsMessage;
      wsInstance.onopen = () => {
        setWsConnected(true);
        setCurrentLog("WebSocket connected. Engine ready.");
        if (reconnectTimer.current) {
          clearTimeout(reconnectTimer.current);
          reconnectTimer.current = null;
        }
      };
      wsInstance.onclose = () => {
        setWsConnected(false);
        setCurrentLog("WebSocket disconnected. Reconnecting in 3s...");
        reconnectTimer.current = setTimeout(connectWebSocket, 3000);
      };
      wsInstance.onerror = (err) => {
        console.error("WebSocket error:", err);
        setCurrentLog("WebSocket connection error.");
      };
      ws.current = wsInstance;
    } catch (err) {
      console.error("Failed to create WebSocket:", err);
      reconnectTimer.current = setTimeout(connectWebSocket, 3000);
    }
  }, [handleWsMessage]);

  useEffect(() => {
    const link = document.createElement("link");
    link.href =
      "https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap";
    link.rel = "stylesheet";
    document.head.appendChild(link);

    const handleMouseMove = (e) => {
      const pattern = document.querySelector(".hero-pattern");
      if (pattern) {
        const moveX = (e.clientX / window.innerWidth) * 20;
        const moveY = (e.clientY / window.innerHeight) * 20;
        pattern.style.backgroundPosition = `${moveX}px ${moveY}px`;
      }
    };
    window.addEventListener("mousemove", handleMouseMove);

    connectWebSocket();

    return () => {
      window.removeEventListener("mousemove", handleMouseMove);
      if (reconnectTimer.current) clearTimeout(reconnectTimer.current);
      if (ws.current) ws.current.close();
    };
  }, [connectWebSocket]);

  const processFile = async (file) => {
    if (!file) return;

    if (file.size > MAX_FILE_SIZE_FRONTEND) {
      setFileUploadError("File terlalu besar. Ukuran maksimal: 2MB");
      setCurrentLog("❌ File upload rejected: exceeds 2MB limit");
      setIsProcessingFile(false);
      return;
    }

    setSelectedFile(file);
    setFileUploadError("");
    setExtractedDocText("");
    setIsProcessingFile(true);

    try {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch(`${API_BASE}/upload`, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errData = await response
          .json()
          .catch(() => ({ detail: "Upload failed" }));
        throw new Error(errData.detail || `HTTP ${response.status}`);
      }

      const data = await response.json();

      if (data.success && data.extracted_text) {
        setExtractedDocText(data.extracted_text);
        setCurrentLog(
          `📄 File "${file.name || "pasted_image.png"}" berhasil dibaca. ${data.char_count} karakter diekstrak.`,
        );
      } else {
        setExtractedDocText(
          data.extracted_text || "[Gagal mengekstrak teks dari dokumen]",
        );
        setFileUploadError(
          data.extracted_text || "Gagal membaca konten dokumen.",
        );
        setCurrentLog(
          `⚠️ File "${file.name || "pasted_image.png"}" terbaca tapi teks tidak bisa diekstrak.`,
        );
      }
    } catch (err) {
      console.error("File upload error:", err);
      setFileUploadError(err.message || "Gagal mengunggah file");
      setCurrentLog(`❌ Gagal mengunggah file: ${err.message}`);
    } finally {
      setIsProcessingFile(false);
    }
  };

  const handleFileSelect = (e) => {
    const file = e.target.files[0];
    if (file) processFile(file);
  };

  const removeFile = () => {
    setSelectedFile(null);
    setExtractedDocText("");
    setFileUploadError("");
  };

  const handlePaste = (e) => {
    const items = e.clipboardData?.items;
    if (!items) return;

    for (let i = 0; i < items.length; i++) {
      if (items[i].type.indexOf("image") !== -1) {
        e.preventDefault();
        const file = items[i].getAsFile();
        if (file) {
          const namedFile = new File([file], `pasted_image_${Date.now()}.png`, {
            type: file.type,
          });
          processFile(namedFile);
        }
        return;
      }
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDraggingOver(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDraggingOver(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDraggingOver(false);

    const files = e.dataTransfer?.files;
    if (files && files.length > 0) {
      processFile(files[0]);
    }
  };

  const startProcess = () => {
    if (!input.trim() && !extractedDocText) return;
    setView("discussion");
    setMessages([]);
    setProgress(10);
    setCurrentLog("Initiating quantitative orchestration engine...");
    setDiscussionPhase("roles");
    setActiveAgent(null);
    setSimulationData(null);
    setSimulationResults(null);

    let pText = "";

    if (extractedDocText) {
      const fileTypeInfo = getFileTypeInfo(selectedFile?.name);
      pText += `[ISI DOKUMEN UPLOAD: ${selectedFile?.name || "pasted_image.png"} (${fileTypeInfo.label})]\n`;
      pText += "═══════════════════════════════════════\n";
      pText += extractedDocText;
      pText += "\n═══════════════════════════════════════\n\n";
    }

    if (input.trim()) {
      if (extractedDocText) {
        pText += `[CATATAN TAMBAHAN USER]: ${input.trim()}`;
      } else {
        pText += input.trim();
      }
    }

    setProblemText(pText.substring(0, 100) + (pText.length > 100 ? "..." : ""));

    if (ws.current && ws.current.readyState === WebSocket.OPEN) {
      ws.current.send(JSON.stringify({ type: "start", problem: pText }));
    }
    setInput("");
  };

  const sendAnswer = () => {
    if (!input.trim()) return;
    setIsWaitingUser(false);
    setCurrentLog("User response injected. Resuming discussion pathways...");
    setDiscussionPhase("discussing");
    setActiveAgent(null);
    if (ws.current && ws.current.readyState === WebSocket.OPEN) {
      ws.current.send(JSON.stringify({ type: "answer", text: input }));
    }
    setInput("");
  };

  const handleRunSimulation = (adjustedParams) => {
    if (ws.current && ws.current.readyState === WebSocket.OPEN) {
      ws.current.send(
        JSON.stringify({
          type: "run_simulation",
          adjusted_params: adjustedParams,
        }),
      );
    }
  };

  const handleConfirmSimulation = (simulationSummary) => {
    if (ws.current && ws.current.readyState === WebSocket.OPEN) {
      ws.current.send(
        JSON.stringify({
          type: "confirm_simulation",
          simulation_summary: simulationSummary,
        }),
      );
    }
    setCurrentLog("Simulation confirmed. Compiling Action Plan & SOP...");
  };

  const handleSkipSimulation = () => {
    if (ws.current && ws.current.readyState === WebSocket.OPEN) {
      ws.current.send(JSON.stringify({ type: "skip_simulation" }));
    }
    setCurrentLog(
      "Skipping simulation. Compiling Action Plan from discussion data...",
    );
  };

  const formatBlueprint = (text) => {
    if (!text) return "";
    return text
      .replace(/(Rp [\d.,]+)/g, '<span class="num-highlight">$1</span>')
      .replace(
        /(\d{1,3}(?:\.\d{3})*(?:,\d+)?\s*(?:pieces|orang|jam|hari|bulan|%))/gi,
        '<span class="num-highlight">$1</span>',
      )
      .replace(
        /(\d+\s*[+\-×÷/]\s*\d+\s*=\s*[\d.,]+)/g,
        '<span class="formula-highlight">$1</span>',
      );
  };

  // ==========================================
  // HANDLER KLIK AGENT UNTUK FULL TEXT
  // ==========================================
  const handleAgentClick = (agentName) => {
    const agentMessages = messages.filter((m) => m.agent === agentName);
    const fullText = agentMessages.map(m => m.text).join('\n\n');
    if (fullText) {
      setSelectedAgentText({ name: agentName, text: fullText });
    }
  };

  // ==========================================
  // RENDER
  // ==========================================

  if (accounts.length === 0) {
    return (
      <div className="min-h-screen bg-[#f8f9fa] text-[#191c1d] font-sans antialiased selection:bg-[#e1e0ff] selection:text-[#07006c]">
        <style>{`
          @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
          .hero-pattern { background-image: radial-gradient(circle at 2px 2px, #e1e3e4 1px, transparent 0); background-size: 40px 40px; }
          @keyframes fadeIn { 0% { opacity: 0; transform: translateY(4px); } 100% { opacity: 1; transform: translateY(0); } }
          .animate-fadeIn { animation: fadeIn 0.3s ease-out forwards; }
        `}</style>
        <AuthGuardScreen />
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#f8f9fa] text-[#191c1d] font-sans antialiased selection:bg-[#e1e0ff] selection:text-[#07006c] relative">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
        .hero-pattern { background-image: radial-gradient(circle at 2px 2px, #e1e3e4 1px, transparent 0); background-size: 40px 40px; }
        .glass-panel { backdrop-filter: blur(12px); background: rgba(255, 255, 255, 0.85); }
        .node-pulse { animation: pulse-ring 3s cubic-bezier(0.4, 0, 0.6, 1) infinite; }
        @keyframes pulse-ring { 0% { transform: scale(0.95); opacity: 0.8; } 50% { transform: scale(1.05); opacity: 0.4; } 100% { transform: scale(0.95); opacity: 0.8; } }
        .num-highlight { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; background-color: #eef2ff; color: #4338ca; padding: 1px 4px; border-radius: 4px; font-weight: 600; font-size: 0.95em; }
        .formula-highlight { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; background-color: #f0fdf4; color: #15803d; padding: 2px 6px; border-radius: 4px; font-weight: 600; }
        .phase-indicator { transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1); }
        .agent-card { transition: all 0.3s ease; }
        .agent-card:hover { transform: translateY(-2px); box-shadow: 0 10px 40px -10px rgba(70, 72, 212, 0.15); }
        .progress-glow { box-shadow: 0 0 20px rgba(70, 72, 212, 0.3); }
        input[type="range"] { -webkit-appearance: none; appearance: none; height: 6px; border-radius: 999px; background: #e2e8f0; outline: none; }
        input[type="range"]::-webkit-slider-thumb { -webkit-appearance: none; appearance: none; width: 18px; height: 18px; border-radius: 50%; background: #4648d4; cursor: pointer; border: 2.5px solid white; box-shadow: 0 2px 6px rgba(70, 72, 212, 0.35); transition: transform 0.15s; }
        input[type="range"]::-webkit-slider-thumb:hover { transform: scale(1.15); }
        input[type="range"]::-moz-range-thumb { width: 18px; height: 18px; border-radius: 50%; background: #4648d4; cursor: pointer; border: 2.5px solid white; box-shadow: 0 2px 6px rgba(70, 72, 212, 0.35); }
        @keyframes slideDown { 0% { transform: translate(-50%, -20px); opacity: 0; } 100% { transform: translate(-50%, 0); opacity: 1; } }
        .animate-slideDown { animation: slideDown 0.3s ease-out forwards; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .animate-spin { animation: spin 1s linear infinite; }
        @keyframes fadeIn { 0% { opacity: 0; transform: translateY(4px); } 100% { opacity: 1; transform: translateY(0); } }
        .animate-fadeIn { animation: fadeIn 0.3s ease-out forwards; }
      `}</style>

      {/* ==================== AGENT DISCUSSION MODAL ==================== */}
      {selectedAgentText && (
        <div className="fixed inset-0 bg-black/50 z-[70] flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white rounded-2xl w-full max-w-2xl shadow-2xl flex flex-col max-h-[80vh] animate-fadeIn">
            <div className="flex items-center justify-between p-5 border-b border-slate-100">
              <h3 className="font-bold text-lg text-slate-900 flex items-center gap-2">
                <Bot className="w-6 h-6 text-[#4648d4]" />
                Full Discussion: {selectedAgentText.name}
              </h3>
              <button
                onClick={() => setSelectedAgentText(null)}
                className="p-2 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-6 overflow-y-auto flex-1">
              <p className="text-sm text-slate-700 leading-relaxed whitespace-pre-wrap">
                {selectedAgentText.text}
              </p>
            </div>
          </div>
        </div>
      )}

      {apiError && (
        <div className="fixed inset-0 bg-black/50 z-[60] flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white rounded-2xl p-6 max-w-md w-full shadow-2xl border border-red-100 animate-fadeIn">
            <div className="flex items-center gap-3 mb-4">
              <div className="w-10 h-10 rounded-full bg-red-50 flex items-center justify-center">
                <AlertOctagon className="w-5 h-5 text-red-500" />
              </div>
              <h3 className="text-lg font-bold text-slate-900">System Alert</h3>
            </div>
            <p className="text-sm text-slate-600 leading-relaxed mb-6">
              {ERROR_MESSAGES[apiError] || "Terjadi kesalahan pada sistem. Silakan coba lagi."}
            </p>
            <button
              onClick={() => setApiError(null)}
              className="w-full py-2.5 bg-slate-900 text-white rounded-xl font-semibold text-sm hover:bg-slate-800 transition-colors"
            >
              Tutup
            </button>
          </div>
        </div>
      )}

      {/* NAV */}
      <nav className="bg-white/80 backdrop-blur-xl border-b border-[#e1e3e4] sticky top-0 z-50">
        <div className="flex justify-between items-center w-full px-6 py-4 max-w-screen-2xl mx-auto">
          <div className="flex items-center gap-8">
            <img
              src={logoArgunex}
              alt="Argunex AI Logo"
              className="h-12 w-auto cursor-pointer object-contain transition-transform duration-200 hover:scale-105"
              onClick={() => setView("home")}
            />
            <div className="hidden md:flex gap-6">
              <button
                onClick={() => setView("workspace")}
                className={`font-medium py-1 transition-colors ${view === "workspace" ? "text-[#4648d4] border-b-2 border-[#4648d4]" : "text-[#464554] hover:text-[#4648d4]"}`}
              >
                Workspace
              </button>
            </div>
          </div>
          <div className="flex items-center gap-4">
            {discussionPhase !== "idle" && (
              <div
                className={`flex items-center gap-2 px-3 py-1.5 rounded-full text-xs font-bold border ${getPhaseStyle(discussionPhase)}`}
              >
                {getPhaseIcon(discussionPhase)} {discussionPhase.toUpperCase()}
              </div>
            )}

            <UserProfileDropdown />
          </div>
        </div>
      </nav>

      <main className="relative">
        {/* ==================== HOME ==================== */}
        {view === "home" && (
          <div className="relative min-h-[calc(100vh-72px)] flex flex-col items-center px-6 py-20 md:py-28 hero-pattern">
            <div className="max-w-5xl mx-auto z-10 text-center mb-24">
              <div className="inline-flex items-center gap-2 px-4 py-1.5 bg-[#4648d4]/10 text-[#4648d4] rounded-full mb-8 border border-[#4648d4]/20 shadow-inner">
                <Sparkles className="w-4 h-4 node-pulse" />
                <span className="text-sm font-bold uppercase tracking-widest">
                  Enterprise Architecture Engine
                </span>
              </div>
              <h1 className="text-6xl md:text-7xl font-extrabold text-[#191c1d] mb-8 leading-[1.1] tracking-tighter">
                The Future of{" "}
                <span className="text-[#4648d4]">Autonomous Orchestration</span>
              </h1>
              <p className="text-xl md:text-2xl text-[#464554] mb-14 max-w-3xl mx-auto leading-relaxed font-medium">
                Deploy, visualize, and collaborate with specialized AI agents in
                a unified workspace. Built for developers who demand surgical
                precision and calm authority.
              </p>
              <button
                onClick={() => setView("workspace")}
                className="bg-white hover:bg-gray-50 border border-gray-200 text-[#4648d4] px-12 py-5 rounded-3xl flex items-center gap-4 font-extrabold text-xl shadow-2xl hover:shadow-indigo-100 transition-all transform hover:-translate-y-1 mx-auto group"
              >
                Launch System Engine
                <span className="text-[#4648d4] text-2xl group-hover:translate-x-1.5 transition-transform font-bold">
                  →
                </span>
              </button>
            </div>
            <div className="w-full max-w-6xl grid grid-cols-1 lg:grid-cols-2 gap-10 border-t border-gray-200 pt-16 z-10">
              <div className="bg-white rounded-[32px] p-10 shadow-xl border border-gray-100 flex flex-col justify-between hover:shadow-2xl transition-all">
                <div>
                  <div className="flex items-center gap-5 mb-8">
                    <div className="w-14 h-14 bg-[#4648d4]/10 text-[#4648d4] rounded-2xl flex items-center justify-center border border-[#4648d4]/20">
                      <Brain className="w-7 h-7" />
                    </div>
                    <div>
                      <h3 className="font-extrabold text-gray-950 text-2xl tracking-tight">
                        Core Engine Specifications
                      </h3>
                      <p className="text-sm text-gray-400 font-bold uppercase tracking-wider mt-0.5">
                        Quantitative Accuracy Mode
                      </p>
                    </div>
                  </div>
                  <ul className="space-y-6">
                    <li className="flex gap-4 items-start text-base text-gray-800 leading-relaxed font-medium">
                      <span className="w-3 h-3 rounded-full bg-[#4648d4] mt-2.5 shrink-0" />
                      <div>
                        <strong className="text-gray-950 font-extrabold text-lg block mb-0.5">
                          Chain-of-Thought Enforcement
                        </strong>
                        Setiap agen WAJIB menuliskan rumus matematis sebelum
                        menyebutkan angka final, mencegah halusinasi kalkulasi.
                      </div>
                    </li>
                    <li className="flex gap-4 items-start text-base text-gray-800 leading-relaxed font-medium">
                      <span className="w-3 h-3 rounded-full bg-[#4648d4] mt-2.5 shrink-0" />
                      <div>
                        <strong className="text-gray-950 font-extrabold text-lg block mb-0.5">
                          Fact Grounding Lock
                        </strong>
                        Angka dari prompt user (gaji, kapasitas, jam kerja)
                        dianggap sumber kebenaran mutlak yang dilarang diubah
                        oleh AI.
                      </div>
                    </li>
                    <li className="flex gap-4 items-start text-base text-gray-800 leading-relaxed font-medium">
                      <span className="w-3 h-3 rounded-full bg-[#4648d4] mt-2.5 shrink-0" />
                      <div>
                        <strong className="text-gray-950 font-extrabold text-lg block mb-0.5">
                          Document Intelligence
                        </strong>
                        Membaca dan menganalisis isi dokumen Word, Excel, PDF,
                        dan gambar secara langsung untuk ekstraksi data
                        kuantitatif.
                      </div>
                    </li>
                  </ul>
                </div>
              </div>
              <div className="bg-white rounded-[32px] p-10 shadow-xl border border-gray-100 flex flex-col justify-between hover:shadow-2xl transition-all">
                <div>
                  <div className="flex items-center gap-5 mb-8">
                    <div className="w-14 h-14 bg-emerald-50 text-emerald-600 rounded-2xl flex items-center justify-center border border-emerald-100">
                      <BarChart3 className="w-7 h-7" />
                    </div>
                    <div>
                      <h3 className="font-extrabold text-gray-950 text-2xl tracking-tight">
                        Argunex AI vs Regular Chatbots
                      </h3>
                      <p className="text-sm text-gray-400 font-bold uppercase tracking-wider mt-0.5">
                        Advanced Architecture Power
                      </p>
                    </div>
                  </div>
                  <ul className="space-y-5">
                    <li className="flex gap-4 items-start text-base text-gray-800 leading-relaxed font-medium p-4 bg-gray-50 rounded-2xl border border-gray-100">
                      <Cpu className="w-6 h-6 text-emerald-600 mt-0.5 shrink-0" />
                      <div>
                        <strong className="text-gray-950 font-extrabold text-lg block mb-0.5">
                          Math-Aware Multi-Agent Debate
                        </strong>
                        Bukan opini bebas, melainkan perdebatan terstruktur
                        dengan verifikasi rumus di setiap langkah.
                      </div>
                    </li>
                    <li className="flex gap-4 items-start text-base text-gray-800 leading-relaxed font-medium p-4 bg-gray-50 rounded-2xl border border-gray-100">
                      <Shield className="w-6 h-6 text-emerald-600 mt-0.5 shrink-0" />
                      <div>
                        <strong className="text-gray-950 font-extrabold text-lg block mb-0.5">
                          Human-in-the-Loop Interruption
                        </strong>
                        Moderator hanya bertanya jika ada parameter kuantitatif
                        yang benar-benar hilang, bukan mengulang hal yang sudah
                        jelas.
                      </div>
                    </li>
                    <li className="flex gap-4 items-start text-base text-gray-800 leading-relaxed font-medium p-4 bg-gray-50 rounded-2xl border border-gray-100">
                      <FileText className="w-6 h-6 text-emerald-600 mt-0.5 shrink-0" />
                      <div>
                        <strong className="text-gray-950 font-extrabold text-lg block mb-0.5">
                          Document-Aware Analysis
                        </strong>
                        Bisa membaca isi file Excel, Word, PDF, dan gambar —
                        bahkan via Paste (Ctrl+V) atau Drag & Drop.
                      </div>
                    </li>
                  </ul>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ==================== WORKSPACE ==================== */}
        {view === "workspace" && (
          <div className="flex-grow flex flex-col items-center justify-center px-6 py-16 relative">
            <div className="w-full max-w-3xl flex flex-col items-center text-center gap-4 mb-12">
              <h1 className="text-4xl font-bold tracking-tight text-[#191c1d]">
                How can <span className="text-[#4648d4]">Argunex AI</span>{" "}
                assist your workflow today?
              </h1>
              <p className="text-lg text-[#464554]">
                Upload, Paste, or Drag & Drop your documents here.
              </p>
            </div>
            <div className="w-full max-w-3xl relative">
              <div
                className={`bg-white rounded-3xl p-4 flex flex-col gap-4 border shadow-lg transition-all ${isDraggingOver ? "border-[#4648d4] border-2 shadow-[0_0_30px_rgba(70,72,212,0.2)]" : "border-[#e1e3e4]"}`}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
              >
                {isDraggingOver && (
                  <div className="absolute inset-0 bg-[#4648d4]/5 backdrop-blur-sm rounded-3xl z-10 flex flex-col items-center justify-center pointer-events-none">
                    <UploadCloud className="w-16 h-16 text-[#4648d4] mb-4 animate-bounce" />
                    <p className="text-xl font-bold text-[#4648d4]">
                      Drop your file here
                    </p>
                    <p className="text-sm text-slate-500 mt-1">
                      Word, Excel, PDF, or Image
                    </p>
                  </div>
                )}

                <textarea
                  value={input}
                  onChange={(e) => setInput(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter" && !e.shiftKey) {
                      e.preventDefault();
                      startProcess();
                    }
                  }}
                  onPaste={handlePaste}
                  maxLength={MAX_INPUT_LENGTH}
                  className="w-full border-none focus:ring-0 text-base bg-transparent resize-none p-2 placeholder:text-gray-400 outline-none h-32"
                  placeholder="Describe your operational problem... (You can also Ctrl+V to paste an image here)"
                />

                <div className="flex justify-end px-2">
                  <span className={`text-[11px] font-medium ${input.length >= MAX_INPUT_LENGTH ? "text-red-500" : "text-slate-400"}`}>
                    {input.length}/{MAX_INPUT_LENGTH}
                  </span>
                </div>

                {selectedFile && (
                  <div className="mx-2 p-4 bg-slate-50 rounded-2xl border border-slate-200 animate-fadeIn">
                    <div className="flex items-center justify-between mb-2">
                      <div className="flex items-center gap-3">
                        {(() => {
                          const ft = getFileTypeInfo(selectedFile.name);
                          const FtIcon = ft.icon;
                          return <FtIcon className={`w-5 h-5 ${ft.color}`} />;
                        })()}
                        <div>
                          <p className="text-sm font-semibold text-slate-800">
                            {selectedFile.name}
                          </p>
                          <p className="text-[10px] text-slate-500">
                            {getFileTypeInfo(selectedFile.name).label} •{" "}
                            {(selectedFile.size / 1024).toFixed(1)} KB
                          </p>
                        </div>
                      </div>
                      <button
                        onClick={removeFile}
                        className="text-red-500 hover:text-red-700 text-xs font-semibold p-1.5 rounded-lg hover:bg-red-50 transition-colors"
                      >
                        <X className="w-4 h-4" />
                      </button>
                    </div>

                    {isProcessingFile && (
                      <div className="flex items-center gap-2 mt-2 text-sm text-indigo-600">
                        <Loader2 className="w-4 h-4 animate-spin" />
                        <span className="font-medium">
                          Membaca isi dokumen...
                        </span>
                      </div>
                    )}

                    {fileUploadError && !isProcessingFile && (
                      <div className="mt-2 p-2 bg-amber-50 border border-amber-200 rounded-lg text-xs text-amber-700 flex items-start gap-2">
                        <AlertTriangle className="w-3.5 h-3.5 mt-0.5 shrink-0" />
                        <span>{fileUploadError}</span>
                      </div>
                    )}

                    {extractedDocText && !isProcessingFile && (
                      <div className="mt-2">
                        <div className="flex items-center gap-2 mb-1">
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-500" />
                          <span className="text-[10px] font-bold text-emerald-600 uppercase tracking-wider">
                            Dokumen berhasil dibaca •{" "}
                            {extractedDocText.length.toLocaleString()} karakter
                          </span>
                        </div>
                        <details className="mt-1">
                          <summary className="text-xs text-slate-500 cursor-pointer hover:text-[#4648d4] font-medium transition-colors">
                            👁️ Preview isi dokumen yang diekstrak
                          </summary>
                          <div className="mt-2 p-3 bg-white rounded-xl border border-slate-200 max-h-48 overflow-y-auto">
                            <pre className="text-xs text-slate-600 whitespace-pre-wrap font-mono leading-relaxed">
                              {extractedDocText.substring(0, 2000)}
                              {extractedDocText.length > 2000
                                ? "\n\n[... sisa teks dipotong untuk preview ...]"
                                : ""}
                            </pre>
                          </div>
                        </details>
                      </div>
                    )}
                  </div>
                )}

                <div className="flex justify-between items-center px-2">
                  <div className="flex flex-col sm:flex-row sm:items-center gap-3">
                    <label
                      className={`flex items-center gap-2 px-4 py-2.5 ${isProcessingFile ? "bg-slate-100 text-slate-400" : "bg-[#f3f4f5] hover:bg-gray-200 text-[#464554]"} rounded-xl transition-all cursor-pointer text-xs font-bold border border-gray-200 shadow-sm active:scale-95 shrink-0 ${isProcessingFile ? "pointer-events-none" : ""}`}
                    >
                      <span className="text-[18px]">📄</span>
                      <span>
                        {isProcessingFile ? "Membaca..." : "Upload Dokumen"}
                      </span>
                      <input
                        type="file"
                        accept=".png,.jpg,.jpeg,.pdf,.doc,.docx,.xls,.xlsx"
                        className="hidden"
                        onChange={handleFileSelect}
                        disabled={isProcessingFile}
                      />
                    </label>
                    <span className="text-[11px] text-gray-400 font-medium leading-tight">
                      Upload, Paste (Ctrl+V), or Drag & Drop • Word, Excel, PDF,
                      PNG, JPG • Max 2MB
                    </span>
                  </div>
                  <button
                    onClick={startProcess}
                    disabled={
                      isProcessingFile || (!input.trim() && !extractedDocText)
                    }
                    className="bg-[#4648d4] text-white h-11 w-11 rounded-full flex items-center justify-center hover:scale-105 transition-all shadow-md shrink-0 disabled:bg-slate-300 disabled:cursor-not-allowed disabled:hover:scale-100"
                  >
                    <span className="font-bold text-lg">↑</span>
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ==================== DISCUSSION ==================== */}
        {view === "discussion" && (
          <div className="h-[calc(100vh-72px)] flex flex-col p-4 md:p-6 relative bg-slate-50">
            <div className="w-full bg-white rounded-2xl p-4 mb-4 shadow-sm border border-slate-200">
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-2">
                  <Activity className="w-4 h-4 text-[#4648d4]" />
                  <span className="text-xs font-bold text-slate-600 uppercase tracking-wider">
                    Orchestration Progress
                  </span>
                </div>
                <span className="text-xs font-bold text-[#4648d4]">
                  {progress}%
                </span>
              </div>
              <div className="w-full h-2.5 bg-slate-200 rounded-full overflow-hidden">
                <div
                  className="h-full bg-gradient-to-r from-[#4648d4] to-[#6063ee] transition-all duration-700 rounded-full progress-glow"
                  style={{ width: `${progress}%` }}
                />
              </div>
            </div>
            <div className="flex-1 flex gap-4 min-h-0">
              <div className="flex-1 bg-white rounded-3xl shadow-lg border border-slate-200 overflow-hidden relative">
                <div className="absolute top-4 left-4 z-10 flex items-center gap-2 px-3 py-1.5 bg-white/90 backdrop-blur-sm rounded-xl border border-slate-200 shadow-sm">
                  <Network className="w-4 h-4 text-[#4648d4]" />
                  <span className="text-xs font-bold text-slate-700">
                    Multi-Agent Network
                  </span>
                </div>
                <div className="absolute top-4 right-4 z-10">
                  <div
                    className={`phase-indicator flex items-center gap-2 px-4 py-2 rounded-xl border text-xs font-bold shadow-sm ${getPhaseStyle(discussionPhase)}`}
                  >
                    {getPhaseIcon(discussionPhase)}{" "}
                    {discussionPhase.toUpperCase()}
                  </div>
                </div>
                <NetworkGraph
                  zoomLevel={zoomLevel}
                  activeAgent={activeAgent}
                  hoveredNode={hoveredNode}
                  setHoveredNode={setHoveredNode}
                  problemText={problemText}
                  agents={visibleAgents}
                  onNodeClick={handleAgentClick}
                />
                <div className="absolute bottom-6 left-6 z-20 flex flex-col gap-1">
                  <button
                    onClick={() => setZoomLevel((z) => Math.min(z + 0.2, 1.8))}
                    className="w-9 h-9 bg-white/90 rounded-lg border border-slate-200 flex items-center justify-center shadow-sm hover:bg-white text-slate-700"
                  >
                    <Plus size={16} />
                  </button>
                  <button
                    onClick={() => setZoomLevel((z) => Math.max(z - 0.2, 0.5))}
                    className="w-9 h-9 bg-white/90 rounded-lg border border-slate-200 flex items-center justify-center shadow-sm hover:bg-white text-slate-700"
                  >
                    <Minus size={16} />
                  </button>
                  <button
                    onClick={() => setZoomLevel(1)}
                    className="w-9 h-9 bg-white/90 rounded-lg border border-slate-200 flex items-center justify-center shadow-sm hover:bg-white text-slate-700"
                  >
                    <RotateCcw size={14} />
                  </button>
                </div>
              </div>
              <div className="w-80 flex flex-col gap-4">
                <div className="bg-white rounded-2xl p-4 shadow-sm border border-slate-200 flex-1 overflow-y-auto">
                  <h3 className="text-sm font-bold text-slate-800 mb-3 flex items-center gap-2">
                    <Users className="w-4 h-4 text-[#4648d4]" /> Agent Status
                  </h3>
                  <div className="space-y-3">
                    {visibleAgents.length === 0 ? (
                      <div className="flex flex-col items-center justify-center h-full min-h-[120px] text-slate-400">
                        <Activity className="w-5 h-5 animate-pulse mb-2 opacity-60" />
                        <span className="text-xs font-medium italic">
                          Awaiting AI network allocation...
                        </span>
                      </div>
                    ) : (
                      visibleAgents.map((agent) => {
                        const agentMessages = messages.filter(
                          (m) => m.agent === agent.name,
                        );
                        const latestMsg =
                          agentMessages[agentMessages.length - 1];
                        const status =
                          activeAgent === agent.name ? "active" : "completed";

                        const displayText = () => {
                          if (latestMsg && latestMsg.text && latestMsg.text.trim().length > 3) {
                            return latestMsg.text.substring(0, 100) + (latestMsg.text.length > 100 ? "..." : "");
                          }
                          if (status === "active") {
                            return `Processing ${agent.name.toLowerCase()} analysis...`;
                          }
                          return "Analysis completed.";
                        };

                        return (
                          <div
                            key={agent.id}
                            onClick={() => handleAgentClick(agent.name)}
                            className={`agent-card p-3 rounded-xl border cursor-pointer ${getAgentCardStyle(status)}`}
                          >
                            <div className="flex items-center justify-between mb-1">
                              <div className="flex items-center gap-2">
                                <div
                                  className={`w-2 h-2 rounded-full ${getStatusDot(status)}`}
                                />
                                <span className="text-xs font-bold text-slate-700">
                                  {agent.name}
                                </span>
                              </div>
                              {getAgentStatusIcon(status)}
                            </div>
                            <p className="text-[10px] text-slate-500 line-clamp-2 mt-1 leading-relaxed">
                              {displayText()}
                            </p>
                          </div>
                        );
                      })
                    )}
                  </div>
                </div>
                <div className="bg-slate-900 rounded-2xl p-4 shadow-lg border border-slate-800">
                  <div className="flex items-center gap-2 mb-2">
                    <MessageSquare className="w-3.5 h-3.5 text-emerald-400" />
                    <span className="text-[10px] font-bold text-emerald-400 uppercase tracking-wider">
                      System Trace
                    </span>
                  </div>
                  <div className="h-24 overflow-y-auto">
                    <p className="text-xs font-mono text-slate-300 leading-relaxed break-words">
                      {currentLog}
                    </p>
                  </div>
                  <div className="mt-2 flex items-center gap-1.5">
                    <div className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse" />
                    <span className="text-[10px] text-slate-500 font-mono">
                      LIVE
                    </span>
                  </div>
                </div>
              </div>
            </div>

            {isWaitingUser && (
              <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-md z-40 flex items-center justify-center p-4 md:p-8">
                <div className="bg-white w-full max-w-lg rounded-3xl shadow-2xl border border-slate-200 p-6 md:p-8 flex flex-col max-h-[85vh]">
                  <div className="flex items-center gap-4 mb-5">
                    <div className="w-14 h-14 bg-amber-50 text-amber-600 rounded-2xl flex items-center justify-center shadow-lg shadow-amber-100">
                      <AlertTriangle size={28} />
                    </div>
                    <div>
                      <h2 className="text-xl font-extrabold tracking-tight text-slate-950">
                        Moderator Interruption
                      </h2>
                      <p className="text-xs text-slate-500 font-medium">
                        Clarification required to proceed
                      </p>
                    </div>
                  </div>
                  <div className="bg-amber-50/80 border border-amber-100 rounded-2xl p-4 mb-5 overflow-y-auto max-h-40">
                    <p className="text-base text-slate-800 font-semibold leading-relaxed whitespace-normal break-words">
                      {modQuestion && modQuestion.trim().length > 0
                        ? modQuestion
                        : "Moderator memerlukan data kuantitatif tambahan untuk melanjutkan analisis. Mohon berikan parameter yang mungkin belum disebutkan sebelumnya."}
                    </p>
                  </div>
                  <textarea
                    value={input}
                    onChange={(e) => setInput(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter" && !e.shiftKey) {
                        e.preventDefault();
                        sendAnswer();
                      }
                    }}
                    maxLength={MAX_CLARIFICATION_LENGTH}
                    className="w-full p-4 border border-slate-200 rounded-xl mb-2 text-base font-medium outline-none focus:ring-2 focus:ring-[#4648d4] bg-slate-50 resize-none h-28"
                    placeholder="Provide the missing parameter..."
                  />
                  <div className="flex justify-end mb-4">
                    <span className={`text-[11px] font-medium ${input.length >= MAX_CLARIFICATION_LENGTH ? "text-red-500" : "text-slate-400"}`}>
                      {input.length}/{MAX_CLARIFICATION_LENGTH}
                    </span>
                  </div>
                  <button
                    onClick={sendAnswer}
                    className="w-full bg-[#4648d4] text-white py-3.5 rounded-xl font-bold hover:bg-[#6063ee] transition-all shadow-md text-base active:scale-[0.98]"
                  >
                    Submit Response
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ==================== SIMULATION ==================== */}
        {view === "simulation" && (
          <SimulationView
            simulationData={simulationData}
            simulationResults={simulationResults}
            onRunSimulation={handleRunSimulation}
            onConfirmSimulation={handleConfirmSimulation}
            onSkipSimulation={handleSkipSimulation}
            onBackToDiscussion={() => setView("discussion")}
            wsConnected={wsConnected}
          />
        )}

        {/* ==================== SUMMARY ==================== */}
        {view === "summary" && (
          <div className="max-w-7xl mx-auto py-12 px-6 grid grid-cols-1 lg:grid-cols-12 gap-6">
            <div className="lg:col-span-8">
              <div className="bg-white rounded-[24px] p-8 shadow-sm border border-[#e1e3e4]">
                <div className="flex items-center justify-between mb-6 pb-4 border-b border-gray-100">
                  <div className="flex items-center gap-3">
                    <TrendingUp className="w-6 h-6 text-[#4648d4]" />
                    <h1 className="text-3xl font-bold text-slate-900">
                      Final Resolution Blueprint
                    </h1>
                  </div>
                  <button
                    onClick={() => setView("simulation")}
                    className="flex items-center gap-2 text-sm font-medium text-slate-500 hover:text-[#4648d4] transition-colors border border-slate-200 px-4 py-2 rounded-xl hover:bg-slate-50 shrink-0"
                  >
                    <ArrowLeft className="w-4 h-4" /> Back to Simulation
                  </button>
                </div>
                <div
                  className="text-gray-700 leading-relaxed text-base whitespace-pre-wrap"
                  dangerouslySetInnerHTML={{
                    __html: formatBlueprint(result.content).replace(
                      /\n/g,
                      "<br/>",
                    ),
                  }}
                />
              </div>
            </div>

            <div className="lg:col-span-4 flex flex-col gap-6">
              {/* PDF DOWNLOAD CARD */}
              <div className="bg-white rounded-[24px] p-6 border border-[#e1e3e4] shadow-sm">
                <div className="flex items-start gap-4 mb-6">
                  <div className="w-12 h-12 rounded-xl bg-red-50 text-red-500 flex items-center justify-center">
                    <FileText size={24} />
                  </div>
                  <div>
                    <h3 className="font-bold text-base">Analysis_Report.pdf</h3>
                    <p className="text-xs text-gray-400">Ready to download</p>
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <button
                    onClick={() =>
                      setPdfPreviewUrl(`${API_BASE}${result.files.pdf}`)
                    }
                    disabled={!result.files.pdf}
                    className="flex items-center justify-center gap-2 border border-gray-200 py-2.5 rounded-xl text-xs font-semibold hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    <Eye size={14} /> Preview
                  </button>
                  <button
                    onClick={() => forceDownload(`${API_BASE}${result.files.pdf}`, "Analysis_Report.pdf")}
                    disabled={!result.files.pdf}
                    className={`flex items-center justify-center gap-2 py-2.5 rounded-xl text-xs font-semibold text-center ${result.files.pdf ? "bg-[#4648d4] text-white hover:bg-[#3638b0]" : "bg-slate-200 text-slate-400 cursor-not-allowed"}`}
                  >
                    <Download size={14} /> Download
                  </button>
                </div>
              </div>

              {/* PPTX DOWNLOAD CARD */}
              <div className="bg-white rounded-[24px] p-6 border border-[#e1e3e4] shadow-sm">
                <div className="flex items-start gap-4 mb-6">
                  <div className="w-12 h-12 rounded-xl bg-amber-50 text-amber-600 flex items-center justify-center">
                    <Layers size={24} />
                  </div>
                  <div>
                    <h3 className="font-bold text-base">
                      Executive_Slides.pptx
                    </h3>
                    <p className="text-xs text-gray-400">
                      Multi-Slide Presentation Format
                    </p>
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <button
                    onClick={() => {
                      setCurrentSlideIndex(0);
                      setIsPptPreviewOpen(true);
                    }}
                    disabled={!result.files.ppt}
                    className="flex items-center justify-center gap-2 border border-gray-200 py-2.5 rounded-xl text-xs font-semibold hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    <Eye size={14} /> Preview
                  </button>
                  <button
                    onClick={() => forceDownload(`${API_BASE}${result.files.ppt}`, "Executive_Slides.pptx")}
                    disabled={!result.files.ppt}
                    className={`flex items-center justify-center gap-2 py-2.5 rounded-xl text-xs font-semibold text-center ${result.files.ppt ? "bg-[#4648d4] text-white hover:bg-[#3638b0]" : "bg-slate-200 text-slate-400 cursor-not-allowed"}`}
                  >
                    <Download size={14} /> Download
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ==================== PDF PREVIEW ==================== */}
        {pdfPreviewUrl && (
          <div className="fixed inset-0 bg-black/60 z-50 flex justify-center items-center p-6 backdrop-blur-sm">
            <div className="bg-white w-full h-full max-w-4xl rounded-2xl flex flex-col overflow-hidden">
              <div className="p-4 bg-gray-900 text-white flex justify-between items-center">
                <h3 className="font-bold text-sm">Live PDF Report Preview</h3>
                <button
                  onClick={() => setPdfPreviewUrl(null)}
                  className="text-xs font-bold hover:bg-gray-800 px-3 py-1 rounded-lg"
                >
                  Close (X)
                </button>
              </div>
              <iframe
                src={pdfPreviewUrl}
                className="flex-1 w-full h-full bg-gray-100"
              />
            </div>
          </div>
        )}

        {/* ==================== PPT PREVIEW ==================== */}
        {isPptPreviewOpen && (
          <div className="fixed inset-0 bg-black/70 z-50 flex justify-center items-center p-4 backdrop-blur-md">
            <div className="w-full max-w-4xl flex flex-col gap-4">
              <div className="flex justify-between items-center text-white px-2">
                <p className="text-sm font-semibold tracking-wider uppercase opacity-80 flex items-center gap-2">
                  <Layers size={16} className="text-[#4648d4]" /> Executive
                  Slide Presentasi (Local Preview)
                </p>
                <button
                  onClick={() => setIsPptPreviewOpen(false)}
                  className="bg-white/10 hover:bg-white/20 p-2 rounded-full"
                >
                  <X size={20} />
                </button>
              </div>
              <div className="w-full aspect-[16/9] bg-white rounded-3xl shadow-2xl border border-gray-100 p-12 flex flex-col justify-between relative overflow-hidden">
                <div className="absolute top-0 left-0 w-full h-2 bg-gradient-to-r from-[#4648d4] to-[#6063ee]" />
                {pptSlides.length > 0 ? (
                  <>
                    <div>
                      <span className="text-[10px] font-bold text-[#4648d4] uppercase tracking-widest bg-[#4648d4]/10 px-3 py-1 rounded-full">
                        Slide {currentSlideIndex + 1} of {pptSlides.length}
                      </span>
                      <h2 className="text-3xl font-bold text-gray-900 mt-4 tracking-tight">
                        {pptSlides[currentSlideIndex].title}
                      </h2>
                    </div>
                    <div className="flex-1 flex flex-col justify-center my-6">
                      <ul className="space-y-4">
                        {pptSlides[currentSlideIndex].bullets.map(
                          (bullet, bIdx) => (
                            <li
                              key={bIdx}
                              className="text-lg text-gray-700 flex items-start gap-3 leading-relaxed"
                            >
                              <span className="w-2 h-2 rounded-full bg-[#4648d4] mt-3 shrink-0" />
                              <span>{bullet}</span>
                            </li>
                          ),
                        )}
                      </ul>
                    </div>
                    <div className="flex justify-between items-center border-t border-gray-100 pt-4 text-xs text-gray-400 font-medium">
                      <span>Argunex Intelligence Engine</span>
                      <span>Confidential Strategy Blueprint</span>
                    </div>
                  </>
                ) : (
                  <div className="flex-1 flex items-center justify-center text-gray-400 font-medium text-sm">
                    Memuat struktur halaman presentasi...
                  </div>
                )}
              </div>
              <div className="flex justify-between items-center px-4">
                <button
                  disabled={currentSlideIndex === 0}
                  onClick={() =>
                    setCurrentSlideIndex((p) => Math.max(0, p - 1))
                  }
                  className="flex items-center gap-2 px-5 py-2.5 bg-white rounded-xl text-sm font-bold text-gray-700 shadow-md hover:bg-gray-50 disabled:opacity-40 transition-all"
                >
                  <ChevronLeft size={16} /> Previous
                </button>
                <span className="text-sm text-white/80 font-semibold tracking-wider">
                  Slide {currentSlideIndex + 1} / {pptSlides.length}
                </span>
                <button
                  disabled={currentSlideIndex === pptSlides.length - 1}
                  onClick={() =>
                    setCurrentSlideIndex((p) =>
                      Math.min(pptSlides.length - 1, p + 1),
                    )
                  }
                  className="flex items-center gap-2 px-5 py-2.5 bg-[#4648d4] text-white rounded-xl text-sm font-bold shadow-md hover:bg-[#6063ee] disabled:opacity-40 transition-all"
                >
                  Next <ChevronRight size={16} />
                </button>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
