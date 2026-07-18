import React, {
  useState,
  useRef,
  useEffect,
  useMemo,
  useCallback,
} from "react";
import { Html5Qrcode } from "html5-qrcode";
import { AdvantaLogo } from "./AdvantaLogo";
import { UserIcon } from "./UserIcon";
import { User } from "lucide-react";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  PieChart,
  Pie,
  Cell,
  AreaChart,
  Area,
  LineChart,
  Line,
  LabelList,
} from "recharts";


const ORIGINAL_SCRIPT_URL = "";
const SCRIPT_URL = "/api";

const customFetch = async (input: RequestInfo | URL, init?: RequestInit) => {
  const originalFetch = window.fetch;
  let url = typeof input === 'string' ? input : (input instanceof Request ? input.url : '');
  
  // Determine primary target URL (Express Backend /api)
  let primaryUrl = url;
  if (!url.startsWith("http")) {
    if (url.startsWith("?")) {
      primaryUrl = SCRIPT_URL + url;
    } else if (!url.startsWith(SCRIPT_URL)) {
      primaryUrl = SCRIPT_URL + (url.startsWith("/") ? "" : "/") + url;
    } else {
      primaryUrl = url;
    }
  }

  console.log("[API Request]:", primaryUrl, init);

  const res = await originalFetch(primaryUrl, init);
  const contentType = res.headers.get("content-type");
  if (!res.ok || (contentType && contentType.indexOf("application/json") === -1)) {
    throw new Error(`Invalid API response (Status: ${res.status})`);
  }
  return res;
};

const cleanForMatch = (s: any) =>
  String(s || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

const matchNames = (name1: any, name2: any) => {
  const c1 = cleanForMatch(name1);
  const c2 = cleanForMatch(name2);
  if (!c1 || !c2) return false;
  return c1 === c2 || c1.includes(c2) || c2.includes(c1);
};

// Helper function for crop fuzzy matching
const checkCropMatch = (itemCrop: string, filterCrop: string): boolean => {
  if (!filterCrop || filterCrop === "All") return true;
  const ic = String(itemCrop || "")
    .toLowerCase()
    .trim();
  const fc = String(filterCrop || "")
    .toLowerCase()
    .trim();
  return ic === fc;
};

const normalizePosition = (pos: string | undefined): string => {
  if (!pos) return "Unknown";
  const clean = cleanForMatch(pos);
  if (clean === "businessanalyst" || clean === "analyst")
    return "Business Analyst";
  if (clean === "areasalesmanager" || clean === "asm")
    return "Area Sales Manager";
  if (clean === "vegetablessalesmanager" || clean === "vsm")
    return "Vegetables Sales Manager";
  if (clean === "salesmanager" || clean === "sm") return "Sales Manager";
  if (clean === "salesagronomist" || clean === "sa") return "Sales Agronomist";
  if (clean === "businesssolution" || clean === "bs")
    return "Business Solution";
  if (clean === "countryhead") return "Country Head";
  if (clean === "commerciallead") return "Commercial Lead";

  // Custom casing logic for presentation
  if (clean.includes("businessanalyst")) return "Business Analyst";
  if (clean.includes("areasalesmanager") || clean.includes("asm"))
    return "Area Sales Manager";
  if (clean.includes("vegetablessalesmanager"))
    return "Vegetables Sales Manager";
  if (clean.includes("salesmanager") || clean.includes("sm"))
    return "Sales Manager";
  if (
    clean.includes("salesagronomist") ||
    clean.includes("sa") ||
    clean.includes("agronomist")
  )
    return "Sales Agronomist";
  if (clean.includes("businesssolution") || clean.includes("bs"))
    return "Business Solution";
  return pos;
};

const CustomBarBackground = (props: any) => {
  const { x, y, width, height, index, data, activeKey } = props;
  const entry = data && data[index];
  const isActive = entry && activeKey === entry.name;
  if (!isActive) return null;
  return (
    <rect
      x={x - 4}
      y={y}
      width={width + 8}
      height={height}
      fill="rgba(249, 115, 22, 0.15)"
      rx={6}
    />
  );
};

const getPositionRank = (pos: string | undefined): number => {
  // Hirarki Posisi (dari yang tertinggi ke terendah):
  const norm = normalizePosition(pos);
  if (norm === "Country Head") return 1;
  if (norm === "Commercial Lead") return 1;
  if (norm === "Business Analyst") return 1;
  if (norm === "Vegetables Sales Manager") return 2;
  if (norm === "Sales Manager") return 2;
  if (norm === "Area Sales Manager") return 3;
  if (norm === "Sales Agronomist") return 4;
  if (norm === "Business Solution") return 5;
  
  const normLower = norm.toLowerCase();
  if (
    normLower.includes("head") || 
    normLower.includes("director") || 
    normLower.includes("vp") || 
    normLower.includes("lead") ||
    normLower.includes("business analyst")
  ) {
    return 1;
  }
  if (normLower.includes("manager")) return 2;
  return 5;
};

const parseLevelStr = (val: string | number | undefined | null): number => {
  if (val === undefined || val === null || val === "") return NaN;
  if (typeof val === "number") return val;
  const str = String(val).toUpperCase().trim();
  if (str === "ADMIN") return 4;
  // Prioritize explicit digit
  const match = str.match(/\d+/);
  if (match) return parseInt(match[0], 10);
  // Explicit roman numerals as fallback
  if (str.includes("IV")) return 4;
  if (str.includes("III")) return 3;
  if (str.includes("II")) return 2;
  if (str.includes("V")) return 5;
  if (str.includes("I")) return 1;
  return NaN;
};

const getFromRecord = <T,>(
  record: Record<string, T>,
  key: string | undefined,
): T | undefined => {
  if (!key) return undefined;
  const cleanKey = cleanForMatch(key);
  let foundKey = Object.keys(record).find((k) => cleanForMatch(k) === cleanKey);
  if (!foundKey) {
    foundKey = Object.keys(record).find((k) => matchNames(k, key));
  }
  return foundKey ? record[foundKey] : undefined;
};

const getMemberLevel = (
  name: string,
  teamLevels: Record<string, number>,
  teamPositions: Record<string, string>,
  userData: any,
): number => {
  const cleanName = cleanForMatch(name);

  // Check if it's the logged-in user:
  if (cleanName === cleanForMatch(userData?.name)) {
    if (
      userData?.level !== undefined &&
      userData.level !== null &&
      String(userData.level).trim() !== ""
    ) {
      const parsed = parseLevelStr(userData.level);
      if (!isNaN(parsed)) return parsed;
    }
  }

  // Check teamLevels state:
  const lvl = getFromRecord(teamLevels, name);
  if (lvl !== undefined && lvl !== null && !isNaN(lvl)) {
    return lvl;
  }



  // Fallback based on position name:
  const p =
    getFromRecord(teamPositions, name) ||
    (cleanName === cleanForMatch(userData?.name) ? userData?.position : "");
  const rank = getPositionRank(p);
  if (rank === 1) return 5;
  if (rank === 2) return 4;
  if (rank === 3) return 3;
  if (rank === 4) return 2;
  if (rank === 5) return 1;
  return 0;
};

const compareMembersByLevel = (
  a: string,
  b: string,
  teamLevels: Record<string, number>,
  teamPositions: Record<string, string>,
  userData: any,
): number => {
  const lvlA = getMemberLevel(a, teamLevels, teamPositions, userData);
  const lvlB = getMemberLevel(b, teamLevels, teamPositions, userData);
  if (lvlA !== lvlB) {
    return lvlB - lvlA; // Higher level (most senior) first
  }
  return a.localeCompare(b);
};

const getUplineInTeam = (
  member: string,
  teamMembers: string[],
  teamUpLines: Record<string, string>,
): string | null => {
  const cleanMember = cleanForMatch(member);

  // 1. Follow the direct upline path recursively to find the first ancestor who is in active teamMembers
  let currentUpline = getFromRecord(teamUpLines, member);
  const seen = new Set<string>([cleanMember]);
  while (currentUpline && currentUpline.trim() !== "") {
    const cleanUp = cleanForMatch(currentUpline);
    if (seen.has(cleanUp)) break; // Prevents circular loops
    seen.add(cleanUp);

    let found = teamMembers.find((tm) => cleanForMatch(tm) === cleanUp);
    if (!found) {
      found = teamMembers.find((tm) => matchNames(tm, currentUpline));
    }
    if (found) {
      return found; // Direct or transitive manager found from column!
    }
    currentUpline = getFromRecord(teamUpLines, currentUpline);
  }

  // 2. Strict mode: if no direct/transitive upline found in active teamMembers, return null (root)
  return null;
};

const formatNum = (num) => {
  if (!num) return "0";
  const absNum = Math.abs(num);
  let minFrac = 0;
  let maxFrac = 0;

  if (absNum >= 100) {
    minFrac = 0;
    maxFrac = 0;
  } else if (absNum >= 10) {
    minFrac = 1;
    maxFrac = 1;
  } else if (absNum > 0) {
    minFrac = 2;
    maxFrac = 2;
  }

  return Number(num).toLocaleString("en-US", {
    minimumFractionDigits: minFrac,
    maximumFractionDigits: maxFrac,
  });
};

const formatOverviewVal = (
  num: number | undefined | null,
  forceMt?: boolean
): { valueStr: string; unit: string } => {
  const n = num || 0;
  if (n === 0) {
    return { valueStr: "0", unit: forceMt ? "MT" : "Kg" };
  }
  const absVal = Math.abs(n);
  const isMt = forceMt !== undefined ? forceMt : (absVal >= 1000);

  if (isMt) {
    const mtVal = n / 1000;
    const absMtVal = Math.abs(mtVal);
    let dec = 0;
    if (absMtVal < 10) {
      dec = 2;
    } else if (absMtVal < 100) {
      dec = 1;
    } else {
      dec = 0;
    }
    return {
      valueStr: mtVal.toLocaleString("id-ID", {
        minimumFractionDigits: dec,
        maximumFractionDigits: dec,
      }),
      unit: "MT",
    };
  } else {
    const dec = absVal < 10 ? 2 : 1;
    return {
      valueStr: n.toLocaleString("id-ID", {
        minimumFractionDigits: dec,
        maximumFractionDigits: dec,
      }),
      unit: "Kg",
    };
  }
};

const formatOverviewWithUnit = (
  num: number | undefined | null,
  forceMt?: boolean
): string => {
  const formatted = formatOverviewVal(num, forceMt);
  return `${formatted.valueStr} ${formatted.unit}`;
};

const parseTaskDate = (timestamp: any) => {
  if (!timestamp) return null;
  if (timestamp instanceof Date) return timestamp;
  let d = new Date(timestamp);
  if (!isNaN(d.getTime())) return d;

  if (typeof timestamp === "string") {
    const isSlash = timestamp.includes("/");
    const isDash = timestamp.includes("-");
    if (isSlash || isDash) {
      const parts = timestamp.split(/[\s/:-]+/);
      if (parts.length >= 3) {
        if (parts[0].length === 4) {
          const dStr = `${parts[0]}-${parts[1]}-${parts[2]}T${parts[3] || "00"}:${parts[4] || "00"}:${parts[5] || "00"}`;
          d = new Date(dStr);
        } else {
          let year = parts[2];
          if (year.length === 2 && !isNaN(Number(year))) {
            year = "20" + year;
          }
          const dStr = `${year}-${parts[1]}-${parts[0]}T${parts[3] || "00"}:${parts[4] || "00"}:${parts[5] || "00"}`;
          d = new Date(dStr);
        }
      }
    }
  }
  return d && !isNaN(d.getTime()) ? d : null;
};

const depthMapCache = new Map<string, Record<string, number>>();

const buildDepthMap = (
  rootName: string,
  teamProfiles: Record<string, any>,
): Record<string, number> => {
  const cacheKey = `${rootName}_${Object.keys(teamProfiles || {}).length}`;
  if (depthMapCache.has(cacheKey)) {
    return depthMapCache.get(cacheKey)!;
  }
  const depths: Record<string, number> = {};
  const cleanRoot = cleanForMatch(rootName);

  // Find if there's a profile in teamProfiles that matches the rootName by name or by key or email
  let realRootName = rootName;
  const foundProfileKey = Object.keys(teamProfiles).find(
    (k) =>
      cleanForMatch(k) === cleanRoot ||
      (teamProfiles[k]?.email &&
        cleanForMatch(teamProfiles[k].email) === cleanRoot),
  );
  if (foundProfileKey) {
    realRootName = foundProfileKey;
  }
  const cleanRealRoot = cleanForMatch(realRootName);

  depths[cleanRealRoot] = 0;
  if (cleanRoot !== cleanRealRoot) {
    depths[cleanRoot] = 0;
  }

  const queue: string[] = [cleanRealRoot];
  const visited = new Set<string>([cleanRealRoot]);

  while (queue.length > 0) {
    const currentClean = queue.shift()!;
    const currentDepth = depths[currentClean];

    Object.entries(teamProfiles).forEach(([name, p]: [string, any]) => {
      const cleanName = cleanForMatch(name);
      if (cleanName === cleanRealRoot || cleanName === cleanRoot) return;
      const cleanUpline = cleanForMatch(p.upline || "");
      if (cleanUpline === currentClean && !visited.has(cleanName)) {
        visited.add(cleanName);
        depths[cleanName] = currentDepth + 1;
        queue.push(cleanName);
      }
    });
  }

  // For any remaining nodes in teamProfiles, try to traverse up their upline to determine depth.
  Object.entries(teamProfiles).forEach(([name]) => {
    const cleanName = cleanForMatch(name);
    if (depths[cleanName] === undefined) {
      let current = cleanName;
      let climbVisited = new Set<string>();
      let path: string[] = [];
      while (
        current &&
        current !== cleanRealRoot &&
        current !== cleanRoot &&
        !climbVisited.has(current)
      ) {
        climbVisited.add(current);
        path.push(current);
        const currentProfile = Object.values(teamProfiles).find(
          (prof: any) => cleanForMatch(prof.name) === current,
        ) as any;
        if (currentProfile && currentProfile.upline) {
          current = cleanForMatch(currentProfile.upline);
        } else {
          break;
        }
      }
      if (current === cleanRealRoot || current === cleanRoot) {
        for (let i = 0; i < path.length; i++) {
          const node = path[i];
          depths[node] = path.length - i;
        }
      } else {
        depths[cleanName] = 99;
      }
    }
  });

  depthMapCache.set(cacheKey, depths);
  return depths;
};

const getDdaOfUserCache = new Map<string, string>();

const getDdaOfUser = (
  picName: string,
  rootName: string | undefined,
  teamProfiles: Record<string, any> | undefined,
): string => {
  if (!rootName || !teamProfiles) return picName;
  const cleanPic = cleanForMatch(picName);
  const cleanRoot = cleanForMatch(rootName);
  if (!cleanPic || cleanPic === "unknown") return picName;

  const cacheKey = `${cleanPic}_${cleanRoot}_${Object.keys(teamProfiles).length}`;
  if (getDdaOfUserCache.has(cacheKey)) {
    return getDdaOfUserCache.get(cacheKey)!;
  }

  const calculate = (): string => {
    let realRootName = rootName;
    const foundProfileKey = Object.keys(teamProfiles).find(
      (k) =>
        cleanForMatch(k) === cleanRoot ||
        (teamProfiles[k]?.email &&
          cleanForMatch(teamProfiles[k].email) === cleanRoot),
    );
    if (foundProfileKey) {
      realRootName = foundProfileKey;
    }
    const cleanRealRoot = cleanForMatch(realRootName);

    const rootProfile = Object.values(teamProfiles).find(
      (p: any) => cleanForMatch(p.name) === cleanRealRoot,
    ) as any;
    const rootPos = rootProfile?.position || "";
    const rootLevelClean = rootProfile?.level ? String(rootProfile.level).toLowerCase().trim() : "";
    const isBusinessAnalyst =
      cleanForMatch(rootPos) === "businessanalyst" ||
      cleanRealRoot === "adityawiratama" ||
      cleanRoot === "adityawiratama" ||
      cleanRealRoot === "aditya" ||
      cleanRoot === "aditya" ||
      rootLevelClean === "admin";

    if (isBusinessAnalyst) {
      return picName;
    }

    const maxThreshold = 5;

    const depths = { ...buildDepthMap(realRootName, teamProfiles) };

    // Also tag depths for the direct root alias so it functions correctly
    depths[cleanRoot] = 0;

    if ((depths[cleanPic] ?? 99) <= maxThreshold) {
      const matched = Object.keys(teamProfiles).find(
        (k) => cleanForMatch(k) === cleanPic,
      );
      return (matched && teamProfiles[matched]?.name) || picName;
    }

    let current = cleanPic;
    let visited = new Set<string>();

    while (current && current !== cleanRealRoot && current !== cleanRoot) {
      if (visited.has(current)) break;
      visited.add(current);

      const profile = Object.values(teamProfiles).find(
        (p: any) => cleanForMatch(p.name) === current,
      ) as any;
      if (!profile || !profile.upline) break;

      const parentClean = cleanForMatch(profile.upline);
      const parentDepth = depths[parentClean] ?? 99;

      if (parentDepth <= maxThreshold) {
        const matched = Object.keys(teamProfiles).find(
          (k) => cleanForMatch(k) === parentClean,
        );
        return (matched && teamProfiles[matched]?.name) || profile.upline;
      }

      current = parentClean;
    }

    return picName;
  };

  const result = calculate();
  getDdaOfUserCache.set(cacheKey, result);
  return result;
};

const EditModal = ({ isOpen, onClose, item, onSave, isSaving, allHybrids = [] }) => {
  const [newQty, setNewQty] = useState("");

  useEffect(() => {
    if (item) {
      setNewQty(item.stock);
    }
  }, [item]);

  if (!isOpen) return null;

  // Calculate prev month's quantity
  const monthsKeys = [
    "jan", "feb", "mar", "apr", "mei", "jun",
    "jul", "ags", "sep", "okt", "nov", "des",
  ];
  const currentMonthIdx = new Date().getMonth();
  const prevMonthIdx = currentMonthIdx === 0 ? 11 : currentMonthIdx - 1;
  const prevMonthKey = monthsKeys[prevMonthIdx];
  const prevMonthLabel = prevMonthKey.toUpperCase();
  const prevVal = item && item[prevMonthKey] !== undefined ? Number(item[prevMonthKey]) : 0;

  const handleSaveClick = () => {
    onSave(item.id, newQty, {});
  };

  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-4 overflow-y-auto">
      <div className="bg-white w-full max-w-[380px] rounded-[24px] p-6 shadow-2xl border border-[#edecff] animate-in fade-in zoom-in-95 duration-200 my-auto">
        <div className="size-12 bg-[#edecff] rounded-full flex items-center justify-center text-primary mb-4">
          <span className="material-symbols-outlined text-[24px]">
            edit_note
          </span>
        </div>
        <h2 className="text-lg font-semibold text-[#181a2c] leading-tight mb-1">
          Edit Detail LOT & Stock
        </h2>
        <p className="text-[10px] text-[#8E94B7] font-semibold uppercase tracking-wider mb-5">
          Sesuaikan Qty untuk LOT ini
        </p>

        <div className="space-y-4 mb-6">
          {/* Hybrid / Varietas info block */}
          <div className="p-4 bg-[#fbf8ff] border border-[#edecff] rounded-2xl space-y-3">
            <div className="flex justify-between items-center gap-2">
              <span className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider">
                Hybrid / Varietas
              </span>
              <span className="text-xs font-bold text-primary text-right max-w-[180px] truncate" title={item?.hybrid}>
                {item?.hybrid}
              </span>
            </div>
            
            <div className="h-px bg-[#edecff]" />

            <div className="flex justify-between items-center">
              <span className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider">
                Nomor Lot
              </span>
              <span className="text-xs font-mono font-bold text-[#181a2c]">
                {item?.lot}
              </span>
            </div>

            <div className="h-px bg-[#edecff]" />

            <div className="flex justify-between items-center">
              <span className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider">
                Qty Bulan Lalu ({prevMonthLabel})
              </span>
              <span className="text-xs font-bold text-slate-700">
                {prevVal} Kg
              </span>
            </div>
          </div>

          {/* Quantity Input */}
          <div className="space-y-1">
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 block">
              Quantity Baru (Kg) <span className="text-red-500">*</span>
            </label>
            <input
              type="number"
              value={newQty}
              onChange={(e) => setNewQty(e.target.value)}
              placeholder="Masukkan quantity baru"
              className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all"
              required
              autoFocus
            />
          </div>
        </div>

        <div className="flex gap-3">
          <button
            onClick={onClose}
            className="flex-1 py-3 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors"
          >
            Batal
          </button>
          <button
            onClick={handleSaveClick}
            disabled={isSaving || !newQty || isNaN(Number(newQty))}
            className={`flex-[2] py-3 rounded-full font-semibold text-xs uppercase tracking-wider transition-all ${
              isSaving || !newQty || isNaN(Number(newQty))
                ? "bg-[#e0e0fa] text-[#8E94B7] cursor-not-allowed"
                : "bg-gradient-to-r from-primary to-cyan-400 text-white shadow-[0_4px_12px_rgba(21,75,226,0.25)] active:scale-[0.98]"
            }`}
          >
            {isSaving ? "Saving..." : "Simpan"}
          </button>
        </div>
      </div>
    </div>
  );
};

const ConfirmModal = ({ isOpen, onClose, onConfirm, isProcessing }) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-6">
      <div className="bg-white w-full max-w-[340px] rounded-[24px] p-8 shadow-2xl border border-[#edecff] text-center animate-in fade-in zoom-in-95 duration-200">
        <div className="size-16 bg-red-50 rounded-full flex items-center justify-center text-[#ba1a1a] mx-auto mb-5">
          <span className="material-symbols-outlined text-[32px]">
            priority_high
          </span>
        </div>
        <h2 className="text-xl font-semibold text-[#181a2c] leading-tight mb-1">
          Konfirmasi Hapus
        </h2>
        <p className="text-xs font-semibold text-[#8E94B7] uppercase tracking-wider mb-6">
          Anda yakin ingin menghapus LOT ini?
        </p>
        <div className="space-y-2">
          <button
            onClick={onConfirm}
            disabled={isProcessing}
            className="w-full py-3 bg-gradient-to-r from-[#ba1a1a] to-rose-600 text-white rounded-full font-semibold text-xs uppercase tracking-wider active:scale-[0.98] shadow-md"
          >
            {isProcessing ? "Processing..." : "Ya, Hapus Data"}
          </button>
          <button
            onClick={onClose}
            className="w-full py-3 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors"
          >
            Batal
          </button>
        </div>
      </div>
    </div>
  );
};

const LogoutConfirmModal = ({ isOpen, onClose, onConfirm }) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-6">
      <div className="bg-white w-full max-w-[340px] rounded-[24px] p-8 shadow-2xl border border-[#edecff] text-center animate-in fade-in zoom-in-95 duration-200">
        <div className="size-16 bg-red-50 rounded-full flex items-center justify-center text-[#ba1a1a] mx-auto mb-5">
          <span className="material-symbols-outlined text-[32px] text-red-500">
            logout
          </span>
        </div>
        <h2 className="text-xl font-semibold text-[#181a2c] leading-tight mb-1">
          Konfirmasi Keluar
        </h2>
        <p className="text-xs font-semibold text-[#8E94B7] uppercase tracking-wider mb-6">
          Apakah Anda yakin ingin keluar dari aplikasi?
        </p>
        <div className="space-y-2">
          <button
            onClick={onConfirm}
            className="w-full py-3 bg-gradient-to-r from-red-600 to-rose-600 text-white rounded-full font-semibold text-xs uppercase tracking-wider active:scale-[0.98] shadow-md cursor-pointer"
          >
            Ya, Keluar
          </button>
          <button
            onClick={onClose}
            className="w-full py-3 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors cursor-pointer"
          >
            Batal
          </button>
        </div>
      </div>
    </div>
  );
};

const playBeep = () => {
  try {
    const ctx = new (
      window.AudioContext || (window as any).webkitAudioContext
    )();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = "sine";
    osc.frequency.setValueAtTime(1200, ctx.currentTime);
    gain.gain.setValueAtTime(0.2, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.12);
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.start();
    osc.stop(ctx.currentTime + 0.12);
  } catch (err) {
    console.warn("Could not play scan beep:", err);
  }
};

const QrScanModal = ({ isOpen, onClose, onScanSuccess }) => {
  const [scannerError, setScannerError] = useState<string | null>(null);
  const scannerRef = useRef<any>(null);

  useEffect(() => {
    if (!isOpen) {
      setScannerError(null);
      return;
    }

    let isScanningActive = false;
    let html5QrCode: any = null;

    const startCamera = async () => {
      try {
        html5QrCode = new Html5Qrcode("qr-camera-stream");
        scannerRef.current = html5QrCode;

        await html5QrCode.start(
          { facingMode: "environment" },
          {
            fps: 15,
            qrbox: (width, height) => {
              const size = Math.min(width, height) * 0.95;
              return { width: size, height: size };
            },
          },
          (decodedText) => {
            playBeep();
            onScanSuccess(decodedText);
            onClose();
          },
          () => {
            // silent frame error check
          },
        );
        isScanningActive = true;
      } catch (err) {
        console.error("Camera access error:", err);
        setScannerError(
          "Gagal mengakses kamera. Mohon berikan izin kamera pada browser Anda.",
        );
      }
    };

    const timer = setTimeout(() => {
      startCamera();
    }, 250);

    return () => {
      clearTimeout(timer);
      if (html5QrCode) {
        if (isScanningActive) {
          html5QrCode.stop().catch((e) => console.log("Stop failed:", e));
        }
      }
    };
  }, [isOpen, onClose, onScanSuccess]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[120] bg-[#181a2c]/80 backdrop-blur-md flex flex-col justify-center items-center p-6 animate-in fade-in duration-200">
      <style>{`
        #qr-camera-stream {
          background-color: transparent !important;
          border: none !important;
        }
        #qr-camera-stream video {
          object-fit: cover !important;
          width: 100% !important;
          height: 100% !important;
          border-radius: 22px !important;
        }
        #qr-camera-stream canvas {
          display: none !important;
        }
        #qr-camera-stream img {
          display: none !important;
        }
      `}</style>
      <div className="bg-white w-full max-w-[340px] rounded-[32px] overflow-hidden shadow-[0_24px_64px_rgba(24,26,44,0.12)] border border-[#f0edff] p-6 flex flex-col items-center relative animate-in fade-in zoom-in-95 duration-250">
        <button
          onClick={onClose}
          className="absolute top-4 right-4 size-8 rounded-full bg-[#f4f2ff] text-[#8E94B7] hover:text-red-500 hover:bg-red-50 transition-all flex items-center justify-center cursor-pointer z-10"
        >
          <span className="material-symbols-outlined text-lg">close</span>
        </button>

        <h3 className="text-base font-bold text-[#181a2c] tracking-wide mb-1 mt-2 text-center">
          Pindai QR / Barcode
        </h3>
        <p className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider text-center mb-6">
          Arahkan kamera ke kode LOT
        </p>

        {/* Thick elegant gradient container matching the website's theme */}
        <div className="relative w-full aspect-square max-w-[260px] p-[10px] rounded-[44px] bg-gradient-to-br from-primary to-cyan-400 shadow-[0_20px_48px_rgba(21,75,226,0.22)] flex items-center justify-center">
          <div className="w-full h-full bg-white rounded-[34px] p-[12px] flex items-center justify-center relative overflow-hidden">
            <div
              id="qr-camera-stream"
              className="w-full h-full object-cover rounded-[22px] overflow-hidden bg-slate-950"
            ></div>

            {!scannerError && (
              <div className="absolute inset-[16px] pointer-events-none flex items-center justify-center z-10">
                {/* Elegant scanning corners */}
                <div className="absolute top-0 left-0 w-5 h-5 border-t-[3px] border-l-[3px] border-primary rounded-tl-md"></div>
                <div className="absolute top-0 right-0 w-5 h-5 border-t-[3px] border-r-[3px] border-primary rounded-tr-md"></div>
                <div className="absolute bottom-0 left-0 w-5 h-5 border-b-[3px] border-l-[3px] border-cyan-400 rounded-bl-md"></div>
                <div className="absolute bottom-0 right-0 w-5 h-5 border-b-[3px] border-r-[3px] border-cyan-400 rounded-br-md"></div>

                {/* Scanning laser line in matching gradient */}
                <div
                  className="w-[95%] h-0.5 bg-gradient-to-r from-transparent via-primary to-transparent opacity-85 absolute animate-bounce"
                  style={{ top: "15%", animationDuration: "2.5s" }}
                ></div>
              </div>
            )}

            {scannerError && (
              <div className="absolute inset-0 bg-slate-950 text-white flex flex-col items-center justify-center p-4 text-center rounded-[22px] overflow-hidden">
                <span className="material-symbols-outlined text-red-500 text-3xl mb-2">
                  videocam_off
                </span>
                <p className="text-[11px] font-semibold select-none leading-relaxed text-slate-300">
                  {scannerError}
                </p>
              </div>
            )}
          </div>
        </div>

        <div className="mt-6 w-full">
          <button
            type="button"
            onClick={onClose}
            className="w-full h-11 bg-[#f4f2ff] hover:bg-[#edecff] text-[#8E94B7] hover:text-[#181a2c] rounded-full font-bold text-xs uppercase tracking-wider transition-all cursor-pointer"
          >
            Tutup Kamera
          </button>
        </div>
      </div>
    </div>
  );
};

const PartnerEditModal = ({
  isOpen,
  onClose,
  item,
  onSave,
  isSaving,
  activeEmployees,
  allProvinces,
  allCategories,
  userData,
  allGroups,
}) => {
  const [newPic, setNewPic] = useState("");
  const [partnerName, setPartnerName] = useState("");
  const [category, setCategory] = useState("");
  const [province, setProvince] = useState("");
  const [group, setGroup] = useState("");

  const categoriesToDisplay = useMemo(() => {
    const defaultCategories = ["Distributor", "R1", "R2", "RTL"];
    const merged = [...defaultCategories];
    if (allCategories && Array.isArray(allCategories)) {
      allCategories.forEach((cat) => {
        const trimmed = String(cat || "").trim();
        if (
          trimmed &&
          trimmed !== "Uncategorized" &&
          !merged.some((m) => m.toLowerCase() === trimmed.toLowerCase())
        ) {
          merged.push(trimmed);
        }
      });
    }
    return merged;
  }, [allCategories]);

  const groupsToDisplay = useMemo(() => {
    const list = new Set(["Advanta"]);
    if (allGroups && Array.isArray(allGroups)) {
      allGroups.forEach((g) => {
        const trimmed = String(g || "").trim();
        if (trimmed && trimmed !== "-") {
          list.add(trimmed);
        }
      });
    }
    return Array.from(list);
  }, [allGroups]);

  useEffect(() => {
    const userProv = String(userData?.province || userData?.area || "").trim();
    const userGroup = String(userData?.group || "").trim() || "Advanta";

    if (item) {
      const rawPic = String(item.pic || "").trim();
      const cleanRawPic = cleanForMatch(rawPic);
      const matchedPic = activeEmployees?.find((p) => cleanForMatch(p.name) === cleanRawPic)?.name || rawPic;
      setNewPic(matchedPic);
      setPartnerName(String(item.name || "").trim());
      setCategory(String(item.category || "").trim());
      setProvince(String(item.province || "").trim() || userProv);
      setGroup(String(item.group || "").trim() || userGroup);
    } else {
      setNewPic("");
      setPartnerName("");
      setCategory("");
      setProvince(userProv);
      setGroup(userGroup);
    }
  }, [item, activeEmployees, userData]);

  if (!isOpen) return null;

  const isAdd = !!item?.isAdd;

  const handleSave = () => {
    onSave(isAdd ? null : item.id, newPic, {
      isAdd,
      name: partnerName.trim(),
      category: category.trim(),
      province: province.trim(),
      group: group.trim(),
      originalName: isAdd ? "" : (item?.name || ""),
      originalPic: isAdd ? "" : (item?.pic || ""),
      originalProvince: isAdd ? "" : (item?.province || item?.area || ""),
      originalGroup: isAdd ? "" : (item?.group || ""),
      originalCategory: isAdd ? "" : (item?.category || ""),
    });
  };

  const isFormValid =
    partnerName.trim() !== "" &&
    category.trim() !== "" &&
    province.trim() !== "" &&
    group.trim() !== "" &&
    newPic.trim() !== "";

  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-6">
      <div className="bg-white w-full max-w-[380px] rounded-[24px] p-8 shadow-2xl border border-[#edecff] animate-in fade-in zoom-in-95 duration-200">
        <div className="size-14 bg-[#edecff] rounded-full flex items-center justify-center text-primary mb-5">
          <span className="material-symbols-outlined text-[28px]">
            {isAdd ? "add_business" : "manage_accounts"}
          </span>
        </div>
        <h2 className="text-xl font-semibold text-[#181a2c] leading-tight mb-1">
          {isAdd ? "Tambah Partner" : "Edit Partner"}
        </h2>
        <p className="text-[11px] text-[#8E94B7] font-semibold uppercase tracking-wider mb-6">
          {isAdd ? "Buat Partner Baru" : "Sesuaikan Data Partner"}
        </p>

        <div className="space-y-4 mb-6">
          {/* 1. Kiosk Name */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              Nama Partner (Kiosk) <span className="text-red-500">*</span>
            </label>
            <input
              type="text"
              value={partnerName}
              onChange={(e) => setPartnerName(e.target.value)}
              className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] focus:border-primary focus:ring-1 focus:ring-primary/10 rounded-xl px-4 text-xs font-bold text-[#111] outline-none transition-all"
              placeholder="Contoh: Kios Mandiri Tani"
              required
            />
          </div>



          {/* 3. Category */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              Kategori Partner <span className="text-red-500">*</span>
            </label>
            <div className="relative">
              <select
                value={category}
                onChange={(e) => setCategory(e.target.value)}
                className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] focus:border-primary focus:ring-1 focus:ring-primary/10 rounded-xl px-4 text-xs font-bold text-[#111] outline-none transition-all appearance-none pr-10"
                required
              >
                <option value="">-- Pilih Kategori --</option>
                {categoriesToDisplay.map((cat, idx) => (
                  <option key={idx} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
              <span className="material-symbols-outlined absolute right-4 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none text-lg">
                expand_more
              </span>
            </div>
          </div>



          {/* 5. PIC */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              PIC (Karyawan Aktif) <span className="text-red-500">*</span>
            </label>
            <div className="relative">
              <select
                value={newPic}
                onChange={(e) => setNewPic(e.target.value)}
                className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all appearance-none pr-10"
                required
              >
                <option value="">-- Pilih PIC --</option>
                {activeEmployees?.map((emp, idx) => (
                  <option key={idx} value={emp.name}>
                    {emp.name} ({emp.position || "Staff"})
                  </option>
                ))}
              </select>
              <span className="material-symbols-outlined absolute right-4 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none text-lg">
                expand_more
              </span>
            </div>
          </div>
        </div>

        <div className="flex gap-3">
          <button
            onClick={onClose}
            className="flex-1 py-3 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors cursor-pointer"
          >
            Batal
          </button>
          <button
            onClick={handleSave}
            disabled={isSaving || !isFormValid}
            className={`flex-[2] py-3 text-white rounded-full font-semibold text-xs uppercase tracking-wider active:scale-[0.98] transition-all cursor-pointer ${
              isSaving || !isFormValid
                ? "bg-[#e0e0fa] text-[#8E94B7] cursor-not-allowed"
                : "bg-gradient-to-r from-primary to-cyan-400 shadow-[0_4px_12px_rgba(21,75,226,0.25)]"
            }`}
          >
            {isSaving ? "Saving..." : "Simpan"}
          </button>
        </div>
      </div>
    </div>
  );
};

const PartnerDeleteModal = ({
  isOpen,
  onClose,
  onConfirm,
  isProcessing,
  itemName,
}) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-6">
      <div className="bg-white w-full max-w-[340px] rounded-[24px] p-8 shadow-2xl border border-[#edecff] text-center animate-in fade-in zoom-in-95 duration-200">
        <div className="size-16 bg-red-50 rounded-full flex items-center justify-center text-[#ba1a1a] mx-auto mb-5">
          <span className="material-symbols-outlined text-[32px]">
            delete_forever
          </span>
        </div>
        <h2 className="text-xl font-semibold text-[#181a2c] leading-tight mb-1">
          Hapus Partner
        </h2>
        <p className="text-xs font-semibold text-[#181a2c] mb-6">
          Hapus partner{" "}
          <span className="font-bold text-red-700">{itemName}</span> dari
          database?
        </p>
        <div className="space-y-2">
          <button
            onClick={onConfirm}
            disabled={isProcessing}
            className="w-full py-3 bg-gradient-to-r from-[#ba1a1a] to-rose-600 text-white rounded-full font-semibold text-xs uppercase tracking-wider active:scale-[0.98] shadow-md"
          >
            {isProcessing ? "Processing..." : "Ya, Hapus"}
          </button>
          <button
            onClick={onClose}
            className="w-full py-3 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors"
          >
            Batal
          </button>
        </div>
      </div>
    </div>
  );
};

const EmployeeEditModal = ({
  isOpen,
  onClose,
  item,
  onSave,
  isSaving,
  allEmployeeNames,
  userData,
  allProvinces,
  accessRules,
}) => {
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [position, setPosition] = useState("");
  const [province, setProvince] = useState("");
  const [password, setPassword] = useState("");
  const [upline, setUpline] = useState("");

  const isAdd = !!item?.isAdd;

  const userLevel = useMemo(() => {
    if (!userData) return 0;
    if (
      userData.level !== undefined &&
      userData.level !== null &&
      String(userData.level).trim() !== ""
    ) {
      const parsed = parseLevelStr(userData.level);
      if (!isNaN(parsed)) return parsed;
    }
    const rank = getPositionRank(userData.position);
    if (rank === 1) return 5;
    if (rank === 2) return 4;
    if (rank === 3) return 3;
    if (rank === 4) return 2;
    if (rank === 5) return 1;
    return 0;
  }, [userData]);

  const isLoginLevel2 = isAdd && userLevel === 2;

  useEffect(() => {
    if (item) {
      setName(String(item.name || ""));
      setEmail(String(item.user || item.email || ""));
      setPassword(String(item.password || ""));
      setUpline(String(item.upline || ""));

      if (isAdd && userLevel === 2) {
        setPosition("Business Solution");
        setProvince(String(userData?.province || "").trim() || "-");
      } else {
        setPosition(String(item.position || ""));
        setProvince(String(item.province || ""));
      }
    }
  }, [item, isAdd, userLevel, userData]);

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!name.trim()) return;
    const rank = getPositionRank(position);
    let levelVal = 0;
    if (rank === 1) levelVal = 5;
    else if (rank === 2) levelVal = 4;
    else if (rank === 3) levelVal = 3;
    else if (rank === 4) levelVal = 2;
    else if (rank === 5) levelVal = 1;
    onSave(
      isAdd ? "" : item?.name || "",
      {
        name,
        email,
        user: email,
        position,
        province,
        password,
        upline: upline || userData?.name,
        level: levelVal,
        group: userData?.group || "Advanta",
      },
      isAdd,
    );
  };

  const isSelf =
    !isAdd &&
    !!(
      userData?.name &&
      item?.name &&
      String(userData.name).toLowerCase().trim() ===
        String(item.name).toLowerCase().trim()
    );

  const loggedInRank = getPositionRank(userData?.position || "");
  const allPos = useMemo(() => {
    const list = Object.keys(accessRules || {});
    // Remove obsolete 'Sales Manager' if it's still lingering in rules
    const filteredList = list.filter(p => p !== "Sales Manager");
    const defaults = [
      "Business Analyst",
      "Vegetables Sales Manager",
      "Area Sales Manager",
      "Sales Agronomist",
      "Business Solution",
    ];
    defaults.forEach(d => {
      if (!filteredList.includes(d)) filteredList.push(d);
    });
    return filteredList;
  }, [accessRules]);

  const positions = allPos.filter((p) => {
    if (item && normalizePosition(item.position) === normalizePosition(p))
      return true;
    return getPositionRank(p) >= loggedInRank;
  });

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-4">
      <div className="bg-white w-full max-w-[400px] max-h-[85vh] rounded-[28px] shadow-2xl border border-[#edecff] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        {/* Header */}
        <div className="p-6 pb-4 border-b border-[#edecff] flex items-center gap-3 bg-white">
          <div className="size-11 bg-primary/10 rounded-full flex items-center justify-center text-primary">
            <span className="material-symbols-outlined text-[22px]">
              {isAdd ? "person_add" : "face"}
            </span>
          </div>
          <div>
            <h2 className="text-base font-bold text-[#181a2c] leading-tight">
              {isAdd ? "Tambah Karyawan Baru" : "Edit Detail Karyawan"}
            </h2>
            <p className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider mt-0.5">
              {isAdd
                ? "Anggota Tim"
                : `${item?.name} ${isSelf ? "(Anda)" : ""}`}
            </p>
          </div>
        </div>

        {/* Scrollable Form */}
        <div className="flex-1 overflow-y-auto p-6 space-y-4 custom-scrollbar bg-white">
          {/* Nama */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              Nama Lengkap
            </label>
            <input
              type="text"
              value={name}
              onChange={(e) => setName(e.target.value)}
              required
              className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] focus:border-primary focus:ring-1 focus:ring-primary/10 rounded-full px-4 text-xs font-semibold text-[#111] outline-none transition-all"
              placeholder="Masukkan nama"
            />
          </div>

          {/* Username */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              Username
            </label>
            <input
              type="text"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] focus:border-primary focus:ring-1 focus:ring-primary/10 rounded-full px-4 text-xs font-semibold text-[#111] outline-none transition-all"
              placeholder="Masukkan username"
            />
          </div>

          {/* Jabatan / Posisi */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              Posisi / Jabatan
            </label>
            <div className="relative">
              <select
                value={position}
                onChange={(e) => setPosition(e.target.value)}
                disabled={isSelf || isLoginLevel2}
                className={`w-full h-11 border border-[#edecff] focus:border-primary focus:ring-1 focus:ring-primary/10 rounded-full px-4 text-xs font-semibold outline-none transition-all appearance-none pr-10 ${isSelf || isLoginLevel2 ? "bg-slate-100 text-[#8E94B7] cursor-not-allowed opacity-80" : "bg-[#fbf8ff] text-[#111]"}`}
              >
                <option value="">-- Pilih Posisi --</option>
                {positions.map((p) => (
                  <option key={p} value={p}>
                    {p}
                  </option>
                ))}
              </select>
              <span className="material-symbols-outlined absolute right-4 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none text-lg">
                expand_more
              </span>
            </div>
          </div>

          {/* Provinsi */}
          {!isLoginLevel2 && (
            <div>
              <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
                Provinsi
              </label>
              <div className="relative">
                <select
                  value={province}
                  onChange={(e) => setProvince(e.target.value)}
                  disabled={isSelf}
                  className={`w-full h-11 border border-[#edecff] focus:border-primary focus:ring-1 focus:ring-primary/10 rounded-full px-4 text-xs font-semibold outline-none transition-all appearance-none pr-10 ${isSelf ? "bg-slate-100 text-[#8E94B7] cursor-not-allowed opacity-80" : "bg-[#fbf8ff] text-[#111]"}`}
                >
                  <option value="">-- Pilih Provinsi --</option>
                  {(allProvinces || []).map((p) => (
                    <option key={p} value={p}>
                      {p}
                    </option>
                  ))}
                </select>
                <span className="material-symbols-outlined absolute right-4 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none text-lg">
                  expand_more
                </span>
              </div>
            </div>
          )}

          {/* Password */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              Password Akun
            </label>
            <input
              type="text"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] focus:border-primary focus:ring-1 focus:ring-primary/10 rounded-full px-4 text-xs font-semibold text-[#111] outline-none transition-all"
              placeholder="Password login"
            />
          </div>
        </div>

        {/* Action Buttons */}
        <div className="p-6 border-t border-[#edecff] flex gap-3 bg-slate-50/50">
          <button
            type="button"
            onClick={onClose}
            className="flex-1 h-11 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors cursor-pointer"
          >
            Batal
          </button>
          <button
            type="button"
            onClick={handleSubmit}
            disabled={isSaving}
            className="flex-[2] h-11 bg-gradient-to-r from-primary to-cyan-400 text-white rounded-full font-semibold text-xs uppercase tracking-wider shadow-[0_4px_12px_rgba(21,75,226,0.25)] active:scale-[0.98] transition-all flex items-center justify-center gap-1.5 cursor-pointer"
          >
            {isSaving ? "Saving..." : "Simpan"}
          </button>
        </div>
      </div>
    </div>
  );
};

const EmployeeDeleteModal = ({
  isOpen,
  onClose,
  onConfirm,
  isProcessing,
  itemName,
}) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-6">
      <div className="bg-white w-full max-w-[340px] rounded-[24px] p-8 shadow-2xl border border-[#edecff] text-center animate-in fade-in zoom-in-95 duration-200">
        <div className="size-16 bg-red-50 rounded-full flex items-center justify-center text-[#ba1a1a] mx-auto mb-5">
          <span className="material-symbols-outlined text-[32px]">
            group_remove
          </span>
        </div>
        <h2 className="text-xl font-semibold text-[#181a2c] leading-tight mb-1">
          Hapus Karyawan
        </h2>
        <p className="text-xs font-semibold text-[#181a2c] mb-6">
          Hapus data karyawan{" "}
          <span className="font-bold text-red-700">{itemName}</span> dari
          database?
        </p>
        <div className="space-y-2">
          <button
            onClick={onConfirm}
            disabled={isProcessing}
            className="w-full py-3 bg-gradient-to-r from-[#ba1a1a] to-rose-600 text-white rounded-full font-semibold text-xs uppercase tracking-wider active:scale-[0.98] shadow-md cursor-pointer"
          >
            {isProcessing ? "Processing..." : "Ya, Hapus"}
          </button>
          <button
            onClick={onClose}
            className="w-full py-3 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors cursor-pointer"
          >
            Batal
          </button>
        </div>
      </div>
    </div>
  );
};

const DetailItemSection = ({
  items,
  onEdit,
  onDelete,
  onUploadActivity,
  isSyncing,
  hasChanges,
  title,
  subtitle,
  category,
}) => {
  const getConditionBadge = (cond) => {
    const condition = String(cond).toLowerCase();
    if (condition === "new" || condition === "baru")
      return (
        <div className="px-2.5 py-0.5 rounded-full bg-cyan-100/60 border border-cyan-200 text-cyan-800 text-[8px] font-bold uppercase tracking-wider">
          BARU
        </div>
      );
    if (condition === "berkurang")
      return (
        <div className="px-2.5 py-0.5 rounded-full bg-emerald-100/60 border border-emerald-200 text-emerald-800 text-[8px] font-bold uppercase tracking-wider">
          BERKURANG
        </div>
      );
    if (condition === "bertambah")
      return (
        <div className="px-2.5 py-0.5 rounded-full bg-[#edecff] border border-[#c4c5d8] text-primary text-[8px] font-bold uppercase tracking-wider">
          BERTAMBAH
        </div>
      );
    if (condition === "habis")
      return (
        <div className="px-2.5 py-0.5 rounded-full bg-orange-100/60 border border-orange-200 text-orange-850 text-[8px] font-bold uppercase tracking-wider">
          AKAN DIHAPUS
        </div>
      );
    return (
      <div className="px-2.5 py-0.5 rounded-full bg-red-100/60 border border-red-200 text-red-800 text-[8px] font-bold uppercase tracking-wider">
        TETAP
      </div>
    );
  };

  return (
    <div className="-mx-5 bg-white overflow-hidden rounded-[48px] pt-6 px-5 pb-6 mb-8 shadow-[0_4px_44px_rgba(24,26,44,0.15)] border border-[#edecff] font-sans mt-8 relative">
      <div className="flex justify-between items-start mb-6">
        <div>
          <div className="flex items-center gap-2 mb-1">
            <div className="size-2 bg-primary rounded-full animate-pulse" />
            <h3 className="text-base font-semibold text-[#181a2c] tracking-tight">
              {title}
            </h3>
          </div>
          <p className="text-[11px] text-[#8E94B7] font-semibold uppercase tracking-wider">
            {subtitle}
          </p>
          {category && (
            <p className="text-[9px] text-primary font-bold uppercase tracking-widest mt-0.5">
              {category}
            </p>
          )}
        </div>
        <div className="bg-white shadow-[0_4px_12px_rgba(21,75,226,0.08)] px-3 py-1.5 rounded-full text-right">
          <p className="text-[8px] font-bold text-[#8E94B7] uppercase leading-none mb-0.5">
            Total LOT
          </p>
          <p className="text-xs font-bold text-primary">{items.length}</p>
        </div>
      </div>

      <div className="space-y-3.5 max-h-[480px] overflow-y-auto px-5 py-5 -mx-5 -my-5 custom-scrollbar">
        {items.length === 0 ? (
          <div className="py-12 flex flex-col items-center justify-center text-center px-4 bg-white shadow-sm rounded-[24px]">
            <p className="text-[11px] font-semibold text-[#8E94B7] uppercase tracking-wider">
              Feed Kosong
            </p>
          </div>
        ) : (
          items.map((item, index) => (
            <div
              key={item.id || index}
              className="group bg-[#fbfaff] rounded-[18px] p-3.5 flex flex-col gap-3 shadow-[0_16px_36px_rgba(21,75,226,0.25)] hover:shadow-[0_20px_48px_rgba(21,75,226,0.32)] transition-all duration-300"
            >
              <div className="flex items-start justify-between w-full gap-2 font-sans">
                <div className="flex flex-col">
                  <div className="flex items-center gap-2 flex-wrap">
                    <p className="font-semibold text-[#181a2c] text-sm tracking-tight">
                      {item.lot}
                    </p>
                    <div className="px-2 py-0.5 rounded-full bg-red-100/50 border border-red-200 mt-0.5">
                      <p className="text-[8.5px] font-bold text-red-700 uppercase tracking-wide">
                        EXP: {item.expired}
                      </p>
                    </div>
                  </div>
                  {(() => {
                    const monthsKeys = [
                      "jan",
                      "feb",
                      "mar",
                      "apr",
                      "mei",
                      "jun",
                      "jul",
                      "ags",
                      "sep",
                      "okt",
                      "nov",
                      "des",
                    ];
                    const currentMonthIdx = new Date().getMonth();
                    const prevMonthIdx =
                      currentMonthIdx === 0 ? 11 : currentMonthIdx - 1;
                    const prevMonthKey = monthsKeys[prevMonthIdx];
                    const prevVal =
                      item[prevMonthKey] !== undefined
                        ? Number(item[prevMonthKey])
                        : 0;
                    return (
                      <p className="text-[10px] text-[#8E94B7] mt-0.5 font-medium">
                        Bulan lalu ({prevMonthKey.toUpperCase()}):{" "}
                        <span className="text-slate-700 font-semibold">
                          {prevVal} Kg
                        </span>
                      </p>
                    );
                  })()}
                </div>
                <div className="flex-shrink-0">
                  {getConditionBadge(item.condition)}
                </div>
              </div>

              <div className="flex items-center justify-between w-full gap-2">
                <div className="flex items-center gap-2 bg-primary px-3 py-1.5 rounded-full min-w-0 flex-1 shadow-[0_10px_24px_rgba(21,75,226,0.38)]">
                  <p className="text-[11px] text-white font-semibold uppercase tracking-tight whitespace-normal break-words leading-tight flex-1">
                    {item.hybrid}
                  </p>
                  <div className="w-px h-3.5 bg-white/25 flex-shrink-0" />
                  <p className="text-[11px] font-black text-white flex-shrink-0">
                    {(() => {
                      const num = Number(item.stock);
                      if (!isNaN(num) && num < 1) {
                        return num.toFixed(2);
                      }
                      return item.stock;
                    })()}{" "}
                    <span className="text-[8.5px] text-white/75 font-medium">
                      Kg
                    </span>
                  </p>
                  <div className="w-px h-3.5 bg-white/25 flex-shrink-0" />
                  <p className="text-[10px] font-black text-amber-300 uppercase tracking-tight flex-shrink-0">
                    {item.aging} BLN
                  </p>
                </div>

                <div className="flex gap-1.5 flex-shrink-0">
                  <button
                    onClick={() => onEdit(item)}
                    className="size-9 rounded-full bg-[#edecff] text-primary hover:bg-[#e6e6ff] transition-all flex items-center justify-center border border-[#c4c5d8] shadow-sm"
                  >
                    <span className="material-symbols-outlined text-[18px]">
                      edit_square
                    </span>
                  </button>
                  <button
                    onClick={() => onDelete(item)}
                    className="size-9 rounded-full bg-red-50 text-red-500 hover:bg-red-100 transition-all flex items-center justify-center border border-red-100 shadow-sm"
                  >
                    <span className="material-symbols-outlined text-[18px]">
                      delete_sweep
                    </span>
                  </button>
                </div>
              </div>
            </div>
          ))
        )}
      </div>

      <button
        onClick={onUploadActivity}
        disabled={isSyncing || !hasChanges}
        className={`w-full h-14 mt-6 text-white rounded-full font-semibold text-xs uppercase tracking-wider shadow-none transition-all flex items-center justify-center gap-3 active:scale-[0.98] ${
          isSyncing || !hasChanges
            ? "bg-[#e0e0fa] text-[#8E94B7] border border-[#edecff] cursor-not-allowed"
            : "bg-gradient-to-r from-primary to-cyan-400 hover:opacity-95 shadow-[0_8px_20px_rgba(21,75,226,0.25)]"
        }`}
      >
        {isSyncing ? (
          <div className="size-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
        ) : (
          <span className="material-symbols-outlined text-sm">
            cloud_upload
          </span>
        )}
        {isSyncing
          ? "Uploading..."
          : hasChanges
            ? "Upload Activity"
            : "Data Tersinkron"}
      </button>
    </div>
  );
};

const Dashboard = ({
  userData,
  activeTab,
  onLogout,
  onUserSwitch,
  setUserData,
  setActiveTab,
  accessRules,
  setAccessRules,
  overviewMetricFilter,
  setOverviewMetricFilter,
  filterBelowMonth,
  setFilterBelowMonth,
  filterBelowChannel,
  setFilterBelowChannel,
  filterBelowMaterial,
  setFilterBelowMaterial,
  filterBelowTeam,
  setFilterBelowTeam,
  filterBelowArea,
  setFilterBelowArea,
  filterBelowCrop,
  setFilterBelowCrop,
  filterBelowType,
  setFilterBelowType,
}) => {
  const [isDropdownOpen, setIsDropdownOpen] = useState(false);
  const dropdownRef = useRef(null);
  const [hoveredLabel, setHoveredLabel] = useState<string | null>(null);
  const [dismissedTooltipLabel, setDismissedTooltipLabel] = useState<string | null>(null);
  const [activeSubBarKey, setActiveSubBarKey] = useState<string | null>(null);
  const [dismissedSubTooltipLabel, setDismissedSubTooltipLabel] = useState<string | null>(null);
  const [showBudgetBar, setShowBudgetBar] = useState<boolean>(true);
  const [showActualBar, setShowActualBar] = useState<boolean>(true);
  const clickedBarRef = useRef<boolean>(false);

  useEffect(() => {
    const handleGlobalClick = () => {
      if (clickedBarRef.current) {
        clickedBarRef.current = false;
        return;
      }
      if (hoveredLabel) {
        setDismissedTooltipLabel(hoveredLabel);
        setDismissedSubTooltipLabel(hoveredLabel);
      }
    };
    window.addEventListener("click", handleGlobalClick);
    return () => window.removeEventListener("click", handleGlobalClick);
  }, [hoveredLabel]);

  const [searchTerm, setSearchTerm] = useState("");
  const [isUserSwitching, setIsUserSwitching] = useState(false);
  const [historyChartType, setHistoryChartType] = useState<
    "opening" | "ending" | "stockIn" | "idle" | "pog"
  >("pog");

  const handleFullscreen = () => {
    try {
      const docEl = document.documentElement;
      if (
        !document.fullscreenElement &&
        !(document as any).webkitFullscreenElement &&
        !(document as any).mozFullScreenElement &&
        !(document as any).msFullscreenElement
      ) {
        if (docEl.requestFullscreen) {
          docEl.requestFullscreen();
        } else if ((docEl as any).webkitRequestFullscreen) {
          (docEl as any).webkitRequestFullscreen();
        } else if ((docEl as any).mozRequestFullScreen) {
          (docEl as any).mozRequestFullScreen();
        } else if ((docEl as any).msRequestFullscreen) {
          (docEl as any).msRequestFullscreen();
        }
      } else {
        if (document.exitFullscreen) {
          document.exitFullscreen();
        } else if ((document as any).webkitExitFullscreen) {
          (document as any).webkitExitFullscreen();
        } else if ((document as any).mozCancelFullScreen) {
          (document as any).mozCancelFullScreen();
        } else if ((document as any).msExitFullscreen) {
          (document as any).msExitFullscreen();
        }
      }
    } catch (e) {
      console.error("Error toggling fullscreen mode:", e);
    }
  };

  const [kiosks, setKiosks] = useState([]);
  const [workingData, setWorkingData] = useState([]);
  const [rawWorkingData, setRawWorkingData] = useState([]);
  const [drSalesData, setDrSalesData] = useState<any[]>([]);
  const [deletedItems, setDeletedItems] = useState([]);
  const [isSyncing, setIsSyncing] = useState(false);
  const [isChannelsLoading, setIsChannelsLoading] = useState(true);
  const [isFetchingData, setIsFetchingData] = useState(true);
  const [employees, setEmployees] = useState<any[]>([]);
  const [isEmployeesLoading, setIsEmployeesLoading] = useState(false);

  // Propose Activity States
  const [proposeBs, setProposeBs] = useState<string>("Lionel Messi");
  const [proposeCategory, setProposeCategory] = useState<string>("Regular");
  const [proposeActivity, setProposeActivity] = useState<string>("Farmer meeting");
  const [generatedProjects, setGeneratedProjects] = useState<any[]>([]);
  const [planningFilterDistrict, setPlanningFilterDistrict] = useState("");
  const [planningFilterSubDistrict, setPlanningFilterSubDistrict] = useState("");
  const [planningFilterMonth, setPlanningFilterMonth] = useState("");
  const [planningStatusFilter, setPlanningStatusFilter] = useState("All");
  const [processingRows, setProcessingRows] = useState<Record<string, boolean>>({});
  const [proposalsList, setProposalsList] = useState<any[]>(() => {
    try {
      const saved = localStorage.getItem("radar_dg_proposals");
      return saved ? JSON.parse(saved) : [];
    } catch (e) {
      return [];
    }
  });
  const [isSubmittingProposal, setIsSubmittingProposal] = useState<boolean>(false);
  const [proposalSuccessMsg, setProposalSuccessMsg] = useState<string>("");
  const [expandedProposalId, setExpandedProposalId] = useState<string | null>(null);

  const bsEmployeesList = useMemo(() => {
    if (!employees || employees.length === 0) {
      return ["Lionel Messi", "Ronaldo", "Mbappe", "Yamal"];
    }
    const list = employees
      .filter((emp) => {
        const pos = normalizePosition(emp.position);
        return pos === "Business Solution";
      })
      .map((emp) => emp.name)
      .filter(Boolean);
    
    // Remove duplicates
    const uniqueList = Array.from(new Set(list));
    if (uniqueList.length === 0) {
      return ["Lionel Messi", "Ronaldo", "Mbappe", "Yamal"];
    }
    return uniqueList;
  }, [employees]);

  useEffect(() => {
    if (bsEmployeesList.length > 0 && !bsEmployeesList.includes(proposeBs)) {
      setProposeBs(bsEmployeesList[0]);
    }
  }, [bsEmployeesList, proposeBs]);

  const computedTeamProfiles = useMemo(() => {
    if (!employees || employees.length === 0) return undefined;
    const profiles: Record<string, any> = {};
    employees.forEach((emp) => {
      // Use clean names for keys to ensure successful lookup
      if (emp.name) profiles[cleanForMatch(emp.name)] = emp;
      if (emp.user) profiles[cleanForMatch(emp.user)] = emp;
      if (emp.email) profiles[cleanForMatch(emp.email)] = emp;
    });
    return profiles;
  }, [employees]);
  const [loadProgress, setLoadProgress] = useState(0);
  const [showLoader, setShowLoader] = useState(true);

  useEffect(() => {
    const isLoadingData =
      isFetchingData || isChannelsLoading || isEmployeesLoading;
    if (isLoadingData) {
      setShowLoader(true);
      setLoadProgress(0);
      let currentProgress = 0;
      const interval = setInterval(() => {
        if (currentProgress < 30) {
          currentProgress += Math.floor(Math.random() * 8) + 4;
        } else if (currentProgress < 60) {
          currentProgress += Math.floor(Math.random() * 5) + 2;
        } else if (currentProgress < 85) {
          currentProgress += Math.floor(Math.random() * 3) + 1;
        } else if (currentProgress < 98) {
          currentProgress += Math.random() > 0.6 ? 1 : 0;
        }
        if (currentProgress > 98) currentProgress = 98;
        setLoadProgress(currentProgress);
      }, 150);

      return () => {
        clearInterval(interval);
      };
    } else {
      setLoadProgress(100);
      const timeout = setTimeout(() => {
        setShowLoader(false);
      }, 500);
      return () => clearTimeout(timeout);
    }
  }, [isFetchingData, isChannelsLoading, isEmployeesLoading]);

  // Dimension filter for Executive Overview chart (Area, Province, Sales Agronomist, Hybrid, Activity)
  const [overviewGroupDimension, setOverviewGroupDimension] = useState<
    "area" | "province" | "sales_agronomist" | "hybrid" | "activity" | "business_solution" | "material"
  >("area");

  // Dimension filter for Sub-performance overview chart
  const [subGroupDimension, setSubGroupDimension] = useState<
    "area" | "province" | "sales_agronomist" | "hybrid" | "activity" | "business_solution" | "material"
  >("province");

  // Dimension filter for Partner Segmentation (Area, Province, Sales Agronomist, Hybrid, Activity)
  const [partnerSegmentDimension, setPartnerSegmentDimension] = useState<
    "area" | "province" | "sales_agronomist" | "hybrid" | "activity"
  >("area");

  // Dynamic parameters for Overview Bar Chart
  const isMovement = overviewMetricFilter === "movement";
  const isArea = overviewGroupDimension === "area";
  const isProvince = overviewGroupDimension === "province";

  const currentBarGap = -15;
  const currentBarCategoryGap = "5%";
  const currentMaxBarSize = 85;

  const isBusinessAnalyst = useMemo(() => {
    if (!userData) return false;
    const isBA = (userData.position &&
        cleanForMatch(userData.position) === "businessanalyst") ||
      cleanForMatch(userData.name || "") === "adityawiratama" ||
      cleanForMatch(userData.name || "") === "aditya";
    const isAdmin = userData.level && String(userData.level).toLowerCase().trim() === "admin";
    return isBA || isAdmin;
  }, [userData]);

  // State Tab Home
  const [selectedKiosk, setSelectedKiosk] = useState("Loading Kiosk...");

  const handleChannelClick = (channelName: string) => {
    const matchedKiosk = kiosks.find(
      (k) => cleanForMatch(k.name) === cleanForMatch(channelName),
    );
    if (matchedKiosk) {
      setSelectedKiosk(matchedKiosk.name);
      if (setActiveTab) {
        setActiveTab("partner");
      }
    }
  };

  const renderMaybeChannelName = (name: string, defaultClass = "") => {
    const isChannel = kiosks.some(
      (k) => cleanForMatch(k.name) === cleanForMatch(name),
    );
    if (isChannel) {
      return (
        <span
          onClick={(e) => {
            e.stopPropagation();
            handleChannelClick(name);
          }}
          className={`${defaultClass} hover:underline decoration-primary hover:text-primary transition-all cursor-pointer inline-flex items-center gap-1 font-bold text-primary group`}
          title="Klik untuk input aktivitas partner ini"
        >
          {name}
          <span className="material-symbols-outlined text-[13px] inline opacity-60 group-hover:opacity-100 group-hover:scale-110 transition-all text-primary">
            edit_note
          </span>
        </span>
      );
    }
    return <span className={defaultClass}>{name}</span>;
  };
  const [lotNo, setLotNo] = useState("");
  const [qty, setQty] = useState("");
  const [editModal, setEditModal] = useState({ isOpen: false, item: null });
  const [deleteModal, setDeleteModal] = useState({ isOpen: false, item: null });
  const [isActionLoading, setIsActionLoading] = useState(false);
  const [lotIntel, setLotIntel] = useState(null);
  const [isLotChecking, setIsLotChecking] = useState(false);
  const [isLotNotFound, setIsLotNotFound] = useState(false);
  const [isScannerOpen, setIsScannerOpen] = useState(false);
  const [manualHybrid, setManualHybrid] = useState("");
  const [manualCrop, setManualCrop] = useState("Field Corn");
  const [manualDrDate, setManualDrDate] = useState("");
  const [manualExpDate, setManualExpDate] = useState("");
  const [manualHybridCustom, setManualHybridCustom] = useState("");

  // State Tab Partner
  const [isLogoutModalOpen, setIsLogoutModalOpen] = useState(false);

  const userLevel = useMemo(() => {
    if (!userData) return 0;
    if (
      userData.level !== undefined &&
      userData.level !== null &&
      String(userData.level).trim() !== ""
    ) {
      const parsed = parseLevelStr(userData.level);
      if (!isNaN(parsed)) return parsed;
    }
    const rank = getPositionRank(userData.position);
    if (rank === 1) return 5;
    if (rank === 2) return 4;
    if (rank === 3) return 3;
    if (rank === 4) return 2;
    if (rank === 5) return 1;
    return 0;
  }, [userData]);

  const [partnerSubTab, setPartnerSubTab] = useState(() => {
    const parsedLevel =
      userData?.level !== undefined &&
      userData?.level !== null &&
      String(userData?.level).trim() !== ""
        ? parseLevelStr(userData.level)
        : userData?.position
          ? getPositionRank(userData.position) === 5
            ? 1
            : 0
          : 0;
    return parsedLevel === 1 ? "channel" : "team";
  });

  useEffect(() => {
    if (userLevel === 1) {
      setPartnerSubTab("channel");
    } else {
      setPartnerSubTab("team");
    }
  }, [userData, userLevel]);

  const [mappingPic, setMappingPic] = useState("");
  const [mappingCategory, setMappingCategory] = useState("");
  const [partnerEditModal, setPartnerEditModal] = useState({
    isOpen: false,
    item: null,
  });
  const [partnerDeleteModal, setPartnerDeleteModal] = useState({
    isOpen: false,
    item: null,
  });
  const [channelsRefreshKey, setChannelsRefreshKey] = useState(0);
  const [isChartFocusedModalOpen, setIsChartFocusedModalOpen] = useState(false);
  const [focusedChartType, setFocusedChartType] = useState<"main" | "sub">("main");
  const [activeMainBarKey, setActiveMainBarKey] = useState<string | null>(null);
  const [activeActivityFilter, setActiveActivityFilter] = useState<string | null>(null);
  const [conversionSalesFilter, setConversionSalesFilter] = useState("activity");
  const [conversionActivityFilter, setConversionActivityFilter] = useState("activity");
  const [activeConversionActivity, setActiveConversionActivity] = useState("Farmer Meeting");
  const [isMonthDropdownOpen, setIsMonthDropdownOpen] = useState(false);
  const monthDropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (monthDropdownRef.current && !monthDropdownRef.current.contains(event.target as Node)) {
        setIsMonthDropdownOpen(false);
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);

  // State Employee Modifying & Hirarki Expand/Collapse
  const [employeeEditModal, setEmployeeEditModal] = useState<{
    isOpen: boolean;
    item: any | null;
  }>({ isOpen: false, item: null });
  const [employeeDeleteModal, setEmployeeDeleteModal] = useState<{
    isOpen: boolean;
    item: any | null;
  }>({ isOpen: false, item: null });
  const [employeesRefreshKey, setEmployeesRefreshKey] = useState(0);



  const activeEmployees = useMemo(() => {
    return (employees || []).filter((emp: any) => {
      const status = String(emp.status || "").trim().toLowerCase();
      // If no status is specified, treat as active. Only filter out if explicit inactive status.
      return status === "aktif" || status === "active" || status === "" || status === "-";
    });
  }, [employees]);


  const [collapsedNodes, setCollapsedNodes] = useState<Record<string, boolean>>(
    {},
  );

  const availableProvinces = useMemo(() => {
    const list = new Set<string>();
    employees.forEach((emp) => {
      const prov = String(emp.province || "").trim();
      if (prov && prov !== "-") {
        list.add(prov);
      }
    });
    // Fallback defaults
    ["Jawa Timur", "Jawa Tengah", "Jawa Barat"].forEach((p) => list.add(p));
    return Array.from(list).sort();
  }, [employees]);

  const availableGroups = useMemo(() => {
    const list = new Set<string>();
    employees.forEach((emp) => {
      const g = String(emp.group || "").trim();
      if (g && g !== "-") {
        list.add(g);
      }
    });
    if (userData?.group) {
      list.add(String(userData.group).trim());
    }
    ["Advanta"].forEach((g) => list.add(g));
    return Array.from(list).sort();
  }, [employees, userData]);

  // State Tab Summary
  const [summaryGroupBy, setSummaryGroupBy] = useState("hybrid"); // Default changed to 'hybrid'
  const [summarySubGroupBy, setSummarySubGroupBy] = useState("channel"); // Sub category
  const [grandTotalViewBy, setGrandTotalViewBy] = useState<"hybrid" | "area">("hybrid");
  const [isSummaryFilterOpen, setIsSummaryFilterOpen] = useState(true);
  const [isOverviewFilterOpen, setIsOverviewFilterOpen] = useState(false);
  const enrichedSummaryDataRef = useRef<any[]>([]);

  // Access Rules
  const [isSavingAccess, setIsSavingAccess] = useState(false);
  const [accessSaveSuccess, setAccessSaveSuccess] = useState(false);

  const handleSaveAccessRules = async () => {
    setIsSavingAccess(true);
    const filteredRules: Record<string, Record<string, boolean>> = {};
    allPositionsList.forEach(pos => {
      if (accessRules[pos]) {
        filteredRules[pos] = {
          home: !!accessRules[pos].home,
          partner: !!accessRules[pos].partner,
          stock: !!accessRules[pos].stock,
          pog: !!accessRules[pos].pog,
          overview: !!accessRules[pos].overview,
          temp: !!accessRules[pos].temp,
          access: !!accessRules[pos].access,
        };
      } else {
        filteredRules[pos] = {
          home: true,
          partner: true,
          stock: true,
          pog: true,
          overview: pos === "Business Analyst",
          temp: pos === "Business Analyst",
          access: pos === "Business Analyst"
        };
      }
    });

    try {
      localStorage.setItem('appAccessRules', JSON.stringify(filteredRules));
      setAccessRules(filteredRules);
    } catch (e) {
      console.error('Failed to save access rules', e);
    }
    try {
      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({
          action: "saveAccessRules",
          rules: filteredRules
        })
      });

      const contentType = resp.headers.get("content-type");
      if (!resp.ok || !contentType || !contentType.includes("application/json")) {
        throw new Error("Respon server tidak valid. Pastikan backend sudah terkonfigurasi.");
      }

      const res = await resp.json();
      if (res.status === "success") {
        setAccessSaveSuccess(true);
      } else {
        console.error("Failed to save access rules to spreadsheet", res.message);
      }
    } catch (err) {
      console.error("Error saving access rules to spreadsheet:", err);
    } finally {
      setIsSavingAccess(false);
      setTimeout(() => setAccessSaveSuccess(false), 3000);
    }
  };

  const allPositionsList = useMemo(() => {
    const list = Object.keys(accessRules || {});
    // Clean up obsolete 'Sales Manager'
    const filteredList = list.filter(p => p !== "Sales Manager");
    const defaults = [
      "Business Analyst",
      "Vegetables Sales Manager",
      "Area Sales Manager",
      "Sales Agronomist",
      "Business Solution"
    ];
    defaults.forEach(d => {
      if (!filteredList.includes(d)) filteredList.push(d);
    });
    return filteredList;
  }, [accessRules]);

  const removeAccessRule = (position: string) => {
    setAccessRules((prev: Record<string, Record<string, boolean>>) => {
      const currentRules = { ...prev };
      delete currentRules[position];
      return currentRules;
    });
  };

  const toggleAccessRule = (position: string, page: string) => {
    setAccessRules((prev: Record<string, Record<string, boolean>>) => {
      const currentRules = prev[position] || {
        home: true,
        partner: true,
        stock: true,
        pog: true,
        overview: false,
        temp: false,
        access: false
      };
      return {
        ...prev,
        [position]: {
          ...currentRules,
          [page]: !currentRules[page]
        }
      };
    });
  };

  const renderAccessCheckbox = (position: string, page: string) => {
    const isChecked = !!accessRules[position]?.[page];
    return (
      <button 
        onClick={() => toggleAccessRule(position, page)}
        className={`inline-flex items-center gap-1 text-[10px] font-bold px-2 py-1 rounded-full border uppercase tracking-wide transition-colors ${
          isChecked 
            ? "text-emerald-600 bg-emerald-50 border-emerald-200 hover:bg-emerald-100 cursor-pointer" 
            : "text-slate-400 bg-slate-50 border-slate-200 hover:bg-slate-100 cursor-pointer"
        }`}
      >
        <span className="material-symbols-outlined text-[12px]">
          {isChecked ? "check_circle" : "cancel"}
        </span> 
        {isChecked ? "Yes" : "No"}
      </button>
    );
  };

  // State Tab Temp (Temporary Review Page)
  const [tempSearchQuery, setTempSearchQuery] = useState("");
  const [tempSortBy, setTempSortBy] = useState("checker");
  const [tempSortOrder, setTempSortOrder] = useState<"asc" | "desc">("asc");
  const [expandedTempRowId, setExpandedTempRowId] = useState<string | null>(
    null,
  );
  const [isTempProceeded, setIsTempProceeded] = useState(false);
  const [isConsolidatingDb, setIsConsolidatingDb] = useState(false);
  const [consolidationSuccessMsg, setConsolidationSuccessMsg] = useState("");

  // State Tab POG
  const [pogGroupBy, setPogGroupBy] = useState("hybrid");
  const [pogSubGroupBy, setPogSubGroupBy] = useState("channel");
  const [pogExpandedRows, setPogExpandedRows] = useState({});
  const [isPogFilterOpen, setIsPogFilterOpen] = useState(true);

  const [expandedRows, setExpandedRows] = useState({});
  const CLUSTER_CONFIG = useMemo(
    () => [
      {
        key: "0-2",
        label: "0-2",
        colorHeader: "text-blue-500",
        colorCell: "text-blue-600",
        colorGrand: "text-blue-600",
        colorChild: "text-blue-500",
      },
      {
        key: "2-4",
        label: "2-4",
        colorHeader: "text-amber-500",
        colorCell: "text-amber-500",
        colorGrand: "text-amber-600",
        colorChild: "text-amber-400",
      },
      {
        key: "4-6",
        label: "4-6",
        colorHeader: "text-amber-600",
        colorCell: "text-amber-600",
        colorGrand: "text-amber-700",
        colorChild: "text-amber-500",
      },
      {
        key: "6-9",
        label: "6-9",
        colorHeader: "text-orange-500",
        colorCell: "text-orange-500",
        colorGrand: "text-orange-600",
        colorChild: "text-orange-400",
      },
      {
        key: "9-12",
        label: "9-12",
        colorHeader: "text-red-500",
        colorCell: "text-red-500",
        colorGrand: "text-red-600",
        colorChild: "text-red-400",
      },
      {
        key: ">12",
        label: ">12",
        colorHeader: "text-red-700",
        colorCell: "text-red-700",
        colorGrand: "text-red-800",
        colorChild: "text-red-600",
      },
      {
        key: "Uncategorized",
        label: "N/A",
        colorHeader: "text-slate-500",
        colorCell: "text-slate-600",
        colorGrand: "text-slate-600",
        colorChild: "text-slate-400",
      },
    ],
    [],
  );

  const ALL_CLUSTER_KEYS = useMemo(
    () => CLUSTER_CONFIG.map((c) => c.key),
    [CLUSTER_CONFIG],
  );
  const [selectedClusters, setSelectedClusters] = useState(ALL_CLUSTER_KEYS);

  const toggleRow = (name) =>
    setExpandedRows((prev) => ({ ...prev, [name]: !prev[name] }));
  const togglePogRow = (name) =>
    setPogExpandedRows((prev) => ({ ...prev, [name]: !prev[name] }));

  const teamMembers = useMemo(() => {
    if (!userData) return [];

    const myNameClean = cleanForMatch(userData.name || "");
    const isAdmin = userData.level && String(userData.level).toLowerCase().trim() === "admin";
    const isBusinessAnalyst =
      (userData.position &&
        cleanForMatch(userData.position) === "businessanalyst") ||
      cleanForMatch(userData.name) === "adityawiratama" ||
      cleanForMatch(userData.name) === "aditya" ||
      isAdmin;

    let rawList: string[] = [];

    if (employees && employees.length > 0) {
      if (isBusinessAnalyst) {
        let filteredEmployees = employees;
        if (isAdmin && userData.group) {
          const myGroupClean = cleanForMatch(userData.group);
          if (myGroupClean !== "all" && myGroupClean !== "") {
            filteredEmployees = employees.filter((e) => {
              const empGroupClean = cleanForMatch(e.group || "");
              return empGroupClean === myGroupClean || cleanForMatch(e.name) === myNameClean;
            });
          }
        }
        const allNames = filteredEmployees.map((e) => e.name).filter(Boolean);
        if (!allNames.some((n) => cleanForMatch(n) === myNameClean)) {
          allNames.unshift(userData.name);
        }
        rawList = allNames;
      } else {
        const resultList = [];
        const queue = [myNameClean];
        const visited = new Set(queue);

        while (queue.length > 0) {
          const curr = queue.shift()!;

          const matchingEmp = employees.find(
            (e) =>
              cleanForMatch(e.name) === curr ||
              (e.email && cleanForMatch(e.email) === curr),
          );
          const identifiersToMatch = new Set<string>();
          identifiersToMatch.add(curr);
          if (matchingEmp) {
            if (
              !resultList.some(
                (r) => cleanForMatch(r) === cleanForMatch(matchingEmp.name),
              )
            ) {
              resultList.push(matchingEmp.name);
            }
            if (matchingEmp.name)
              identifiersToMatch.add(cleanForMatch(matchingEmp.name));
            if (matchingEmp.email)
              identifiersToMatch.add(cleanForMatch(matchingEmp.email));
          }

          employees.forEach((emp) => {
            const empUplineClean = cleanForMatch(emp.upline);
            if (empUplineClean && identifiersToMatch.has(empUplineClean)) {
              const empClean = cleanForMatch(emp.name);
              if (empClean && !visited.has(empClean)) {
                visited.add(empClean);
                queue.push(empClean);
              }
            }
          });
        }

        const userMatchedEmp = employees.find(
          (e) =>
            cleanForMatch(e.name) === myNameClean ||
            (e.email && cleanForMatch(e.email) === myNameClean),
        );
        const finalUserName = userMatchedEmp
          ? userMatchedEmp.name
          : userData.name;

        if (
          !resultList.some(
            (r) => cleanForMatch(r) === cleanForMatch(finalUserName),
          )
        ) {
          resultList.unshift(finalUserName);
        }

        rawList = resultList;
      }
    } else if (
      computedTeamProfiles &&
      Object.keys(computedTeamProfiles).length > 0
    ) {
      let realRootName = userData.name;
      const foundProfileKey = Object.keys(computedTeamProfiles).find(
        (k) =>
          cleanForMatch(k) === myNameClean ||
          (computedTeamProfiles[k]?.email &&
            cleanForMatch(computedTeamProfiles[k].email) === myNameClean),
      );
      if (foundProfileKey) {
        realRootName = foundProfileKey;
      }
      const cleanRealRoot = cleanForMatch(realRootName);

      const depths = buildDepthMap(realRootName, computedTeamProfiles);
      const maxDepth = 5;

      const filtered = Object.keys(computedTeamProfiles).filter((name) => {
        const d = depths[cleanForMatch(name)] ?? 99;
        return d <= maxDepth;
      });

      const withoutMe = filtered.filter(
        (n) =>
          cleanForMatch(n) !== cleanRealRoot &&
          cleanForMatch(n) !== myNameClean,
      );
      const result = [realRootName, ...withoutMe];
      if (userData.name !== realRootName && !result.includes(userData.name)) {
        result.unshift(userData.name);
      }
      rawList = result;
    } else {
      let uniquePics = [...(userData.subordinates || [])];

      let added = true;
      while (added) {
        added = false;
        kiosks.forEach((k) => {
          const uplineClean = cleanForMatch(k.upline || "");
          const picStr = String(k.pic || "").trim();
          const picClean = cleanForMatch(picStr);

          const isDirectMatch =
            !isBusinessAnalyst &&
            uplineClean !== "" &&
            myNameClean !== "" &&
            (uplineClean.includes(myNameClean) ||
              myNameClean.includes(uplineClean));
          const isTransitiveMatch =
            uplineClean !== "" &&
            uniquePics.some((m) => cleanForMatch(m) === uplineClean);

          if (isDirectMatch || isTransitiveMatch) {
            if (
              picClean !== "" &&
              picClean !== myNameClean &&
              !picClean.includes(myNameClean)
            ) {
              if (
                !uniquePics.some(
                  (existing) => cleanForMatch(existing) === picClean,
                )
              ) {
                uniquePics.push(picStr);
                added = true;
              }
            }
          }
        });
      }

      // FALLBACK AMAN: Jika hasil filter tim sangat sedikit, atau user menggunakan akun demo/tidak terpetakan,
      // maka masukkan semua PIC yang ada di daftar Kiosks agar data Stock Summary/POG tetap tampil dan dapat dianalisis.
      if (uniquePics.length <= (userData.subordinates?.length || 0)) {
        kiosks.forEach((k) => {
          const picStr = String(k.pic || "").trim();
          const picClean = cleanForMatch(picStr);
          if (picStr && picClean !== "" && picClean !== myNameClean) {
            if (
              !uniquePics.some(
                (existing) => cleanForMatch(existing) === picClean,
              )
            ) {
              uniquePics.push(picStr);
            }
          }
        });
      }

      if (
        !uniquePics.some((existing) => cleanForMatch(existing) === myNameClean)
      ) {
        uniquePics.unshift(userData.name);
      }
      rawList = uniquePics;
    }

    // Apply strict filtering for Level 5: if the current user level is 5, exclude other Level 5 users.
    const isMyLevel5 = (() => {
      const cleanName = cleanForMatch(userData.name);
      if (
        userData.position &&
        cleanForMatch(userData.position) === "businesssolution"
      )
        return true;
      if (
        userData.level !== undefined &&
        userData.level !== null &&
        String(userData.level).trim() === "5"
      )
        return true;
      if (employees && employees.length > 0) {
        const emp = employees.find((e) => cleanForMatch(e.name) === cleanName);
        if (emp) {
          if (
            emp.level !== undefined &&
            emp.level !== null &&
            String(emp.level).trim() === "5"
          )
            return true;
          if (cleanForMatch(emp.position) === "businesssolution") return true;
          if (getPositionRank(emp.position) === 5) return true;
        }
      }
      return false;
    })();

    if (isMyLevel5) {
      return rawList.filter((memberName) => {
        const cleanName = cleanForMatch(memberName);
        if (cleanName === myNameClean) return true; // Keep ourselves always

        // Exclude if level is 5 or position is Business Solution
        if (employees && employees.length > 0) {
          const emp = employees.find(
            (e) => cleanForMatch(e.name) === cleanName,
          );
          if (emp) {
            if (
              emp.level !== undefined &&
              emp.level !== null &&
              String(emp.level).trim() === "5"
            )
              return false;
            if (cleanForMatch(emp.position) === "businesssolution")
              return false;
            if (getPositionRank(emp.position) === 5) return false;
          }
        }
        if (computedTeamProfiles) {
          const foundKey = Object.keys(computedTeamProfiles).find(
            (k) => cleanForMatch(k) === cleanName,
          );
          if (foundKey) {
            const prof = computedTeamProfiles[foundKey];
            if (prof?.level !== undefined && String(prof.level).trim() === "5")
              return false;
            if (
              prof?.position &&
              cleanForMatch(prof.position) === "businesssolution"
            )
              return false;
          }
        }
        return true;
      });
    }

    // Deduplicate the list to prevent key collision React errors
    const seen = new Set();
    const uniqueList = [];
    for (const name of rawList) {
      const clean = cleanForMatch(name);
      if (!seen.has(clean)) {
        seen.add(clean);
        uniqueList.push(name);
      }
    }
    return uniqueList;
  }, [
    kiosks,
    userData,
    userData?.name,
    userData?.subordinates,
    userData?.position,
    computedTeamProfiles,
    employees,
  ]);

  const myActiveSubordinates = useMemo(() => {
    return activeEmployees.filter((emp: any) =>
      teamMembers.some((m: string) => cleanForMatch(m) === cleanForMatch(emp.name))
    );
  }, [activeEmployees, teamMembers]);

  const normalizeName = useCallback(
    (nameStr: string) => {
      const trimmed = nameStr.trim();
      if (!trimmed || trimmed === "Unknown") return "Unknown";
      const clean = cleanForMatch(trimmed);

      if (clean === "agusherdianto" || clean.includes("agusherdianto")) {
        return "AGUS HERDIANTO";
      }

      const found = teamMembers.find((t) => cleanForMatch(t) === clean);
      if (found) return found.trim();

      return trimmed;
    },
    [teamMembers],
  );

  const [teamPositions, setTeamPositions] = useState<Record<string, string>>(
    {},
  );
  const [teamAreas, setTeamAreas] = useState<Record<string, string>>({});
  const [teamProvinces, setTeamProvinces] = useState<Record<string, string>>(
    {},
  );
  const [teamSubordinates, setTeamSubordinates] = useState<
    Record<string, string[]>
  >({});
  const [teamUpLines, setTeamUpLines] = useState<Record<string, string>>({});
  const [teamLevels, setTeamLevels] = useState<Record<string, number>>({});

  useEffect(() => {
    if (!userData) return;
    if (employeesRefreshKey === 0) {
      return;
    }
    setIsEmployeesLoading(true);
    customFetch(`${SCRIPT_URL}?action=getEmployees`)
      .then((res) => res.json())
      .then((res) => {
        if (
          res.status === "success" &&
          Array.isArray(res.data) &&
          res.data.length > 0
        ) {
          setEmployees(res.data);
        } else {
          console.warn(
            "getEmployees returned non-success or empty, falling back to offline employees.",
          );
          setEmployees(OFFLINE_EMPLOYEES);
        }
      })
      .catch((err) => {
        console.warn(
          "Error fetching getEmployees, using offline fallbacks:",
          err,
        );
        setEmployees(OFFLINE_EMPLOYEES);
      })
      .finally(() => {
        setIsEmployeesLoading(false);
      });
  }, [userData, employeesRefreshKey]);

  useEffect(() => {
    if (employees && employees.length > 0) {
      const positions: Record<string, string> = {};
      const areas: Record<string, string> = {};
      const provinces: Record<string, string> = {};
      const uplines: Record<string, string> = {};
      const subordinates: Record<string, string[]> = {};
      const levels: Record<string, number> = {};

      employees.forEach((emp) => {
        const name = emp.name;
        positions[name] = normalizePosition(emp.position);
        areas[name] = String(emp.area || "-").trim();
        provinces[name] = String(emp.province || "-").trim();
        uplines[name] = String(emp.upline || "").trim();
        if (
          emp.level !== undefined &&
          emp.level !== null &&
          String(emp.level).trim() !== ""
        ) {
          const parsed = parseLevelStr(emp.level);
          if (!isNaN(parsed)) {
            levels[name] = parsed;
          } else {
            const rank = getPositionRank(emp.position);
            levels[name] =
              rank === 1
                ? 5
                : rank === 2
                  ? 4
                  : rank === 3
                    ? 3
                    : rank === 4
                      ? 2
                      : rank === 5
                        ? 1
                        : 0;
          }
        } else {
          const rank = getPositionRank(emp.position);
          levels[name] =
            rank === 1
              ? 5
              : rank === 2
                ? 4
                : rank === 3
                  ? 3
                  : rank === 4
                    ? 2
                    : rank === 5
                      ? 1
                      : 0;
        }
      });

      employees.forEach((emp) => {
        const name = emp.name;
        const directSubs: string[] = [];
        employees.forEach((item) => {
          if (item.name !== name && item.upline) {
            if (cleanForMatch(item.upline) === cleanForMatch(name)) {
              directSubs.push(item.name);
            }
          }
        });
        subordinates[name] = directSubs;
      });

      setTeamPositions((prev) => ({ ...prev, ...positions }));
      setTeamAreas((prev) => ({ ...prev, ...areas }));
      setTeamProvinces((prev) => ({ ...prev, ...provinces }));
      setTeamUpLines((prev) => ({ ...prev, ...uplines }));
      setTeamSubordinates((prev) => ({ ...prev, ...subordinates }));
      setTeamLevels((prev) => ({ ...prev, ...levels }));
    }
  }, [employees]);

  useEffect(() => {
    if (
      userData &&
      computedTeamProfiles &&
      Object.keys(computedTeamProfiles).length > 0
    ) {
      const positions: Record<string, string> = {};
      const areas: Record<string, string> = {};
      const provinces: Record<string, string> = {};
      const uplines: Record<string, string> = {};
      const subordinates: Record<string, string[]> = {};
      const levels: Record<string, number> = {};

      Object.entries(computedTeamProfiles).forEach(
        ([name, p]: [string, any]) => {
          positions[name] = normalizePosition(p.position);
          areas[name] = String(p.area || "-").trim();
          provinces[name] = String(p.province || "-").trim();
          uplines[name] = String(p.upline || "").trim();
          if (
            p.level !== undefined &&
            p.level !== null &&
            String(p.level).trim() !== ""
          ) {
            const parsed = parseLevelStr(p.level);
            if (!isNaN(parsed)) {
              levels[name] = parsed;
            } else {
              const rank = getPositionRank(p.position);
              levels[name] =
                rank === 1
                  ? 5
                  : rank === 2
                    ? 4
                    : rank === 3
                      ? 3
                      : rank === 4
                        ? 2
                        : rank === 5
                          ? 1
                          : 0;
            }
          } else {
            const rank = getPositionRank(p.position);
            levels[name] =
              rank === 1
                ? 5
                : rank === 2
                  ? 4
                  : rank === 3
                    ? 3
                    : rank === 4
                      ? 2
                      : rank === 5
                        ? 1
                        : 0;
          }
        },
      );

      Object.keys(computedTeamProfiles).forEach((name) => {
        const directSubs: string[] = [];
        Object.entries(computedTeamProfiles).forEach(
          ([otherName, p]: [string, any]) => {
            if (otherName !== name && p.upline) {
              const cleanUp = cleanForMatch(p.upline);
              const cleanMy = cleanForMatch(name);
              if (cleanUp === cleanMy) {
                directSubs.push(otherName);
              }
            }
          },
        );
        subordinates[name] = directSubs;
      });

      setTeamPositions((prev) => ({ ...prev, ...positions }));
      setTeamAreas((prev) => ({ ...prev, ...areas }));
      setTeamProvinces((prev) => ({ ...prev, ...provinces }));
      setTeamUpLines((prev) => ({ ...prev, ...uplines }));
      setTeamSubordinates((prev) => ({ ...prev, ...subordinates }));
      setTeamLevels((prev) => ({ ...prev, ...levels }));
    }
  }, [userData]);

  useEffect(() => {
    const isAdmin = userData.level && String(userData.level).toLowerCase().trim() === "admin";
    const isBusinessAnalyst =
      (userData.position &&
        cleanForMatch(userData.position) === "businessanalyst") ||
      cleanForMatch(userData.name) === "adityawiratama" ||
      cleanForMatch(userData.name) === "aditya" ||
      isAdmin;
    if (isBusinessAnalyst) {
      const others = teamMembers.filter(
        (m) => cleanForMatch(m) !== cleanForMatch(userData.name),
      );
      setTeamSubordinates((prev) => {
        if (JSON.stringify(prev[userData.name]) === JSON.stringify(others))
          return prev;
        return {
          ...prev,
          [userData.name]: others,
        };
      });
      setTeamPositions((prev) => {
        const targetPos = isAdmin ? (userData.position || "Admin") : "Business Analyst";
        if (prev[userData.name] === targetPos) return prev;
        return {
          ...prev,
          [userData.name]: targetPos,
        };
      });
    }
  }, [userData.name, userData.position, teamMembers, teamPositions]);

  const greeting = useMemo(() => {
    const hour = new Date().getHours();
    if (hour >= 4 && hour < 11) {
      return {
        text: "Selamat Pagi",
        imageUrl:
          "https://lh3.googleusercontent.com/d/1AzKb-75MaU9hppqSdy2rS93t0tAPGkGi",
        color: "text-amber-300",
      };
    }
    if (hour >= 11 && hour < 15) {
      return {
        text: "Selamat Siang",
        imageUrl:
          "https://lh3.googleusercontent.com/d/1ZpNkT7R57FppIpyPuTt2w9QtJdIwwuRp",
        color: "text-yellow-300",
      };
    }
    if (hour >= 15 && hour < 19) {
      return {
        text: "Selamat Sore",
        imageUrl:
          "https://lh3.googleusercontent.com/d/12RsJXxDrH7aIAph0AJubB3i4w0gmkxcL",
        color: "text-orange-400",
      };
    }
    return {
      text: "Selamat Malam",
      imageUrl:
        "https://lh3.googleusercontent.com/d/1wzqPdQ5jvw7fOF2X76kM56l9l-4mUcLx",
      color: "text-indigo-200",
    };
  }, []);

  const fetchWorkingData = async (
    presetData?: any[],
    presetDrSales?: any[],
  ) => {
    setIsFetchingData(true);
    setIsSyncing(true);
    try {
      let combinedData = [];
      let combinedDrSales = [];
      let success = false;

      if (presetData && presetDrSales) {
        combinedData = presetData;
        combinedDrSales = presetDrSales;
        success = true;
      } else {
        try {
          const [resp, respDr] = await Promise.all([
            customFetch(
              `${SCRIPT_URL}?action=getWorkingData&user=${encodeURIComponent(userData.name)}`,
            ),
            customFetch(
              `${SCRIPT_URL}?action=getDrSalesData&user=${encodeURIComponent(userData.name)}`,
            ),
          ]);
          const [res, resDr] = await Promise.all([resp.json(), respDr.json()]);

          if (res.status === "success") {
            combinedData = (res.data || []).map((item) => {
              const cropVal =
                item.crops ||
                item.Crops ||
                item.crop ||
                item.Crop ||
                item.CROP ||
                item.CROPS ||
                "Uncategorized Crops";
              const areaVal = item.area || item.Area || item.AREA;
              return { ...item, crops: cropVal, area: areaVal };
            });
            success = true;
          }
          if (resDr.status === "success") {
            combinedDrSales = resDr.data || [];
          }
        } catch (apiErr) {
          console.warn(
            "API working/sales data load failed, falling back offline:",
            apiErr,
          );
        }
      }

      if (!success || combinedData.length === 0) {
        combinedData = OFFLINE_WORKING_DATA;
        combinedDrSales = OFFLINE_DR_SALES;
      }

      setDrSalesData(combinedDrSales);
      setRawWorkingData(combinedData);

      const groupedMap: Record<string, any> = {};

      const parseTimestamp = (ts) => {
        if (!ts) return 0;
        if (typeof ts === "string") {
          if (ts.includes("/")) {
            const parts = ts.split(/[\s/:]+/);
            if (parts.length >= 3) {
              return new Date(
                `${parts[2]}-${parts[1]}-${parts[0]}T${parts[3] || "00"}:${parts[4] || "00"}:${parts[5] || "00"}`,
              ).getTime();
            }
          }
          const d = new Date(ts).getTime();
          return isNaN(d) ? 0 : d;
        }
        const dt = new Date(ts).getTime();
        return isNaN(dt) ? 0 : dt;
      };

      // 1. Lakukan Grouping (Gabungkan LOT & Hybrid yang sama)
      combinedData.forEach((d) => {
        const k = cleanForMatch(d.kiosk);
        const l = cleanForMatch(d.lot);
        const h = cleanForMatch(d.hybrid);
        const key = `${k}_${l}_${h}`;
        if (!groupedMap[key]) {
          groupedMap[key] = { ...d };
        } else {
          // Ambil data dari timestamp terbaru (TIDAK ADA AKUMULASI QTY)
          const timeExisting = parseTimestamp(groupedMap[key].timestamp);
          const timeNew = parseTimestamp(d.timestamp);
          if (timeNew > timeExisting) {
            groupedMap[key] = { ...d };
          }
        }
      });

      // 2. Filter dan format
      const monthsKeys = [
        "jan",
        "feb",
        "mar",
        "apr",
        "mei",
        "jun",
        "jul",
        "ags",
        "sep",
        "okt",
        "nov",
        "des",
      ];
      const currentMonthIdx = new Date().getMonth();
      const currentMonthKey = monthsKeys[currentMonthIdx];
      const prevMonthIdx = currentMonthIdx === 0 ? 11 : currentMonthIdx - 1;
      const prevMonthKey = monthsKeys[prevMonthIdx];

      const enrichedData = Object.values(groupedMap)
        .filter((d) => String(d.condition).trim().toLowerCase() !== "habis") // Filter: Sembunyikan yang habis
        .map((d, index) => {
          const currVal =
            d[currentMonthKey] !== undefined
              ? Number(d[currentMonthKey])
              : Number(d.stock || 0);
          const prevVal =
            d[prevMonthKey] !== undefined ? Number(d[prevMonthKey]) : 0;

          let finalCondition = d.condition || "tetap";
          if (prevVal > 0) {
            if (currVal > prevVal) finalCondition = "bertambah";
            else if (currVal < prevVal) finalCondition = "berkurang";
            else finalCondition = "tetap";
          } else {
            if (currVal > 0) {
              finalCondition =
                d.condition === "new" || d.condition === "baru" || d.isNew
                  ? "new"
                  : "bertambah";
            } else {
              finalCondition = "tetap";
            }
          }

          return {
            ...d,
            id: d.id || `db_${index}_${d.lot}`, // Fallback ID jika tidak ada
            originalStock: d.stock,
            originalUser: d.user || "",
            condition: finalCondition,
            isNew: false,
          };
        });

      setWorkingData(enrichedData);
      setDeletedItems([]);
    } catch (e) {
      console.warn("Error processing working data", e);
    } finally {
      setIsFetchingData(false);
      setIsSyncing(false);
    }
  };

  useEffect(() => {
    const fetchChannels = async () => {
      if (channelsRefreshKey === 0) {
        if (kiosks.length > 0) {
          if (isBusinessAnalyst) {
            setSelectedKiosk("Not Applicable for Business Analyst");
          } else {
            const myKiosksData = kiosks.filter((k) =>
              matchNames(k.pic, userData.name),
            );
            if (myKiosksData.length > 0) setSelectedKiosk(myKiosksData[0].name);
            else {
              if (kiosks.length > 0) setSelectedKiosk(kiosks[0].name);
              else setSelectedKiosk("No Channel Assigned");
            }
          }
        }
        setIsChannelsLoading(false);
        return;
      }
      setIsChannelsLoading(true);
      try {
        let fetchedKiosks = [];
        let success = false;
        try {
          const resp = await customFetch(
            `${SCRIPT_URL}?action=getChannels&user=${encodeURIComponent(userData.name)}`,
          );
          const res = await resp.json();
          if (
            res.status === "success" &&
            Array.isArray(res.data) &&
            res.data.length > 0
          ) {
            fetchedKiosks = res.data;
            success = true;
          }
        } catch (apiErr) {
          console.warn(
            "API channels load failed, falling back offline:",
            apiErr,
          );
        }

        if (!success) {
          fetchedKiosks = OFFLINE_KIOSKS;
        }

        setKiosks(fetchedKiosks);

        if (channelsRefreshKey === 0) {
          if (isBusinessAnalyst) {
            setSelectedKiosk("Not Applicable for Business Analyst");
          } else {
            const myKiosksData = fetchedKiosks.filter((k) =>
              matchNames(k.pic, userData.name),
            );
            if (myKiosksData.length > 0) setSelectedKiosk(myKiosksData[0].name);
            else {
              if (fetchedKiosks.length > 0)
                setSelectedKiosk(fetchedKiosks[0].name);
              else setSelectedKiosk("No Channel Assigned");
            }
          }
        }
      } catch (e) {
        console.warn("Error loading channels", e);
        setKiosks(OFFLINE_KIOSKS);
        if (channelsRefreshKey === 0) {
          if (OFFLINE_KIOSKS.length > 0)
            setSelectedKiosk(OFFLINE_KIOSKS[0].name);
        }
      } finally {
        setIsChannelsLoading(false);
      }
    };
    fetchChannels();
  }, [userData.name, userData.position, channelsRefreshKey, isBusinessAnalyst]);

  useEffect(() => {
    if (!userData) return;

    const loadInitialData = async () => {
      setIsEmployeesLoading(true);
      setIsChannelsLoading(true);
      setIsFetchingData(true);
      setIsSyncing(true);

      try {
        const resp = await customFetch(
          `${SCRIPT_URL}?action=getInitialData&user=${encodeURIComponent(userData.name)}`,
        );
        const res = await resp.json();

        if (res.status === "success" && res.data) {
          // 0. Update User Data if server profile returned
          if (res.data.profile) {
            setUserData((prev) => {
              if (!prev) return res.data.profile;
              const updated = { ...prev, ...res.data.profile };
              
              const isAditya =
                cleanForMatch(updated.name) === "adityawiratama" ||
                cleanForMatch(updated.name) === "aditya" ||
                cleanForMatch(updated.user || "") === "aditya" ||
                cleanForMatch(updated.user || "") === "adityawiratama";
              if (isAditya) {
                updated.position = "Business Analyst";
              } else {
                updated.position = normalizePosition(updated.position);
              }
              
              localStorage.setItem("radar_user_session", JSON.stringify(updated));
              return updated;
            });
          }
          
          // 1. Employees
          if (res.data.employees && res.data.employees.length > 0) {
            setEmployees(res.data.employees);
          } else {
            setEmployees(OFFLINE_EMPLOYEES);
          }

          // 2. Channels
          const fetchedKiosks =
            res.data.channels && res.data.channels.length > 0
              ? res.data.channels
              : OFFLINE_KIOSKS;
          setKiosks(fetchedKiosks);

          if (channelsRefreshKey === 0) {
            if (isBusinessAnalyst) {
              setSelectedKiosk("Not Applicable for Business Analyst");
            } else {
              const myKiosksData = fetchedKiosks.filter((k) =>
                matchNames(k.pic, userData.name),
              );
              if (myKiosksData.length > 0)
                setSelectedKiosk(myKiosksData[0].name);
              else {
                if (fetchedKiosks.length > 0)
                  setSelectedKiosk(fetchedKiosks[0].name);
                else setSelectedKiosk("No Channel Assigned");
              }
            }
          }

          // 3. Working & Sales Data
          const rawWorking =
            res.data.workingData && res.data.workingData.length > 0
              ? res.data.workingData
              : OFFLINE_WORKING_DATA;
          console.log("rawWorking", rawWorking);
          const mappedWorking = rawWorking.map((item) => {
            const cropVal =
              item.crops ||
              item.Crops ||
              item.crop ||
              item.Crop ||
              item.CROP ||
              item.CROPS ||
              "Uncategorized Crops";
            const areaVal = item.area || item.Area || item.AREA;
            return { ...item, crops: cropVal, area: areaVal };
          });

          const drSales =
            res.data.drSalesData && res.data.drSalesData.length > 0
              ? res.data.drSalesData
              : OFFLINE_DR_SALES;

          // Enriches and updates states inside fetchWorkingData
          await fetchWorkingData(mappedWorking, drSales);

          // 4. Access Rules
          if (res.data.accessRules && Object.keys(res.data.accessRules).length > 0) {
            setAccessRules(res.data.accessRules);
            try {
              localStorage.setItem('appAccessRules', JSON.stringify(res.data.accessRules));
            } catch (e) {
              console.error('Failed to save appAccessRules to localStorage', e);
            }
          }
        } else {
          console.warn(
            "getInitialData not supported or empty, doing live individual parallel fetching.",
          );
          await loadIndividualDataFallback();
        }
      } catch (err) {
        console.warn(
          "Failed unified initial data fetch, falling back to parallel individual live fetches:",
          err,
        );
        await loadIndividualDataFallback();
      } finally {
        setIsEmployeesLoading(false);
        setIsChannelsLoading(false);
        setIsFetchingData(false);
        setIsSyncing(false);
      }
    };

    const loadIndividualDataFallback = async () => {
      try {
        const [respEmp, respChan, respWork, respDr, respAccess] = await Promise.all([
          customFetch(`${SCRIPT_URL}?action=getEmployees`),
          customFetch(
            `${SCRIPT_URL}?action=getChannels&user=${encodeURIComponent(userData.name)}`,
          ),
          customFetch(
            `${SCRIPT_URL}?action=getWorkingData&user=${encodeURIComponent(userData.name)}`,
          ),
          customFetch(
            `${SCRIPT_URL}?action=getDrSalesData&user=${encodeURIComponent(userData.name)}`,
          ),
          customFetch(`${SCRIPT_URL}?action=getAccessRules`),
        ]);
        const [resEmp, resChan, resWork, resDr, resAccess] = await Promise.all([
          respEmp.json(),
          respChan.json(),
          respWork.json(),
          respDr.json(),
          respAccess.json(),
        ]);

        if (resEmp.status === "success") {
          setEmployees(resEmp.data || []);
        } else {
          setEmployees(OFFLINE_EMPLOYEES);
        }

        const fetchedKiosks =
          resChan.status === "success" ? resChan.data || [] : OFFLINE_KIOSKS;
        setKiosks(fetchedKiosks);

        if (channelsRefreshKey === 0) {
          if (isBusinessAnalyst) {
            setSelectedKiosk("Not Applicable for Business Analyst");
          } else {
            const myKiosksData = fetchedKiosks.filter((k) =>
              matchNames(k.pic, userData.name),
            );
            if (myKiosksData.length > 0) setSelectedKiosk(myKiosksData[0].name);
            else {
              if (fetchedKiosks.length > 0)
                setSelectedKiosk(fetchedKiosks[0].name);
              else setSelectedKiosk("No Channel Assigned");
            }
          }
        }

        let mappedWorking = OFFLINE_WORKING_DATA;
        if (resWork.status === "success") {
          mappedWorking = (resWork.data || []).map((item) => {
            const cropVal =
              item.crops ||
              item.Crops ||
              item.crop ||
              item.Crop ||
              item.CROP ||
              item.CROPS ||
              "Uncategorized Crops";
            const areaVal = item.area || item.Area || item.AREA;
            return { ...item, crops: cropVal, area: areaVal };
          });
        }
        const drSales =
          resDr.status === "success" ? resDr.data || [] : OFFLINE_DR_SALES;

        await fetchWorkingData(mappedWorking, drSales);

        if (resAccess.status === "success" && resAccess.data && Object.keys(resAccess.data).length > 0) {
          setAccessRules(resAccess.data);
          try {
            localStorage.setItem('appAccessRules', JSON.stringify(resAccess.data));
          } catch (e) {
            console.error('Failed to save appAccessRules to localStorage', e);
          }
        }
      } catch (fallbackErr) {
        console.warn("Fallback live fetches also failed:", fallbackErr);
        setEmployees(OFFLINE_EMPLOYEES);
        setKiosks(OFFLINE_KIOSKS);
        await fetchWorkingData(OFFLINE_WORKING_DATA, OFFLINE_DR_SALES);
      }
    };

    loadInitialData();
  }, [userData.name]);

  useEffect(() => {
    if (teamMembers.length > 1) {
      if (
        !mappingPic ||
        mappingPic === "" ||
        (mappingPic !== "ALL_TEAM" &&
          !teamMembers.some(
            (m) => cleanForMatch(m) === cleanForMatch(mappingPic),
          ))
      ) {
        setMappingPic("ALL_TEAM");
      }
    } else {
      setMappingPic(userData.name || "");
    }
  }, [userData.name, teamMembers, mappingPic]);

  useEffect(() => {
    setIsLotNotFound(false);
    if (lotNo.length < 3) {
      setLotIntel(null);
      setManualHybrid("");
      setManualHybridCustom("");
      setManualCrop("Field Corn");
      setManualDrDate("");
      setManualExpDate("");
      return;
    }
    const checkLot = async () => {
      setIsLotChecking(true);
      try {
        const resp = await customFetch(
          `${SCRIPT_URL}?action=getLotInfo&lot=${encodeURIComponent(lotNo)}`,
        );
        const res = await resp.json();
        if (res.status === "success" && res.data) {
          setLotIntel(res.data);
          setIsLotNotFound(false);
        } else {
          setLotIntel(null);
          setIsLotNotFound(true);
        }
      } catch (e) {
        setLotIntel(null);
        setIsLotNotFound(true);
      } finally {
        setIsLotChecking(false);
      }
    };
    const timer = setTimeout(checkLot, 600);
    return () => clearTimeout(timer);
  }, [lotNo]);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target))
        setIsDropdownOpen(false);
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const getCurrentTimestamp = () => {
    const pad = (n: number) => String(n).padStart(2, "0");
    const now = new Date();
    return `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  };

  const formatDateToAppFormat = (dateStr: string) => {
    if (!dateStr) return "N/A";
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return "N/A";
    const day = String(date.getDate()).padStart(2, "0");
    const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
    const year = String(date.getFullYear()).substring(2);
    return `${day}/${months[date.getMonth()]}/${year}`;
  };

  const calculateAgingInMonths = (startStr: string) => {
    if (!startStr) return "-";
    const start = new Date(startStr);
    const end = new Date();
    if (isNaN(start.getTime())) return "-";
    return String(Math.round((end.getTime() - start.getTime()) / (1000 * 3600 * 24) / 30.416));
  };

  const handleAddLocal = () => {
    if (!lotNo || !qty) return;
    const cleanLot = lotNo.trim().toUpperCase();
    
    let cleanHybrid = "Unknown";
    let cleanCrops = "";
    let cleanDrDate = "";
    let cleanExpired = "N/A";
    let cleanAging = "-";

    if (lotIntel) {
      cleanHybrid = lotIntel.desc || "Unknown";
      cleanCrops = lotIntel.crops || "";
      cleanDrDate = lotIntel.drDate || "";
      cleanExpired = lotIntel.expDate || "N/A";
      cleanAging = lotIntel.aging || "-";
    } else if (isLotNotFound) {
      const selectedH = manualHybrid === "CUSTOM" ? manualHybridCustom.trim() : manualHybrid.trim();
      cleanHybrid = selectedH || "Unknown";
      cleanCrops = manualCrop;
      cleanDrDate = manualDrDate ? formatDateToAppFormat(manualDrDate) : "N/A";
      cleanExpired = manualExpDate ? formatDateToAppFormat(manualExpDate) : "N/A";
      cleanAging = manualDrDate ? calculateAgingInMonths(manualDrDate) : "-";
    }

    const existingItemIndex = workingData.findIndex(
      (item) =>
        cleanForMatch(item.kiosk) === cleanForMatch(selectedKiosk) &&
        cleanForMatch(item.lot) === cleanForMatch(cleanLot) &&
        cleanForMatch(item.hybrid) === cleanForMatch(cleanHybrid),
    );

    if (existingItemIndex !== -1) {
      setWorkingData((prev) =>
        prev.map((item, index) => {
          if (index === existingItemIndex) {
            const nQty = Number(qty);
            const monthsKeys = [
              "jan",
              "feb",
              "mar",
              "apr",
              "mei",
              "jun",
              "jul",
              "ags",
              "sep",
              "okt",
              "nov",
              "des",
            ];
            const currentMonthIdx = new Date().getMonth();
            const currentMonthKey = monthsKeys[currentMonthIdx];
            const prevMonthIdx =
              currentMonthIdx === 0 ? 11 : currentMonthIdx - 1;
            const prevMonthKey = monthsKeys[prevMonthIdx];
            const prevVal =
              item[prevMonthKey] !== undefined ? Number(item[prevMonthKey]) : 0;

            let cond = "tetap";
            if (prevVal === 0) {
              cond = item.isNew ? "new" : nQty > 0 ? "bertambah" : "tetap";
            } else {
              if (nQty > prevVal) cond = "bertambah";
              else if (nQty < prevVal) cond = "berkurang";
              else cond = "tetap";
            }

            return {
              ...item,
              stock: nQty,
              condition: cond,
              user: userData.name,
              timestamp: getCurrentTimestamp(),
              [currentMonthKey]: nQty,
            };
          }
          return item;
        }),
      );
    } else {
      const monthsKeys = [
        "jan",
        "feb",
        "mar",
        "apr",
        "mei",
        "jun",
        "jul",
        "ags",
        "sep",
        "okt",
        "nov",
        "des",
      ];
      const currentMonthKey = monthsKeys[new Date().getMonth()];
      const newItem = {
        id: "local_" + Date.now(),
        lot: cleanLot,
        hybrid: cleanHybrid,
        crops: cleanCrops,
        drDate: cleanDrDate,
        stock: Number(qty),
        aging: cleanAging,
        expired: cleanExpired,
        kiosk: selectedKiosk,
        condition: "new", // BARU
        isNew: true,
        originalStock: Number(qty),
        user: userData.name,
        timestamp: getCurrentTimestamp(),
        [currentMonthKey]: Number(qty),
      };
      setWorkingData((prev) => [newItem, ...prev]);
    }

    // Reset Form
    setLotNo("");
    setQty("");
    setLotIntel(null);
    setIsLotNotFound(false);
    setManualHybrid("");
    setManualHybridCustom("");
    setManualCrop("Field Corn");
    setManualDrDate("");
    setManualExpDate("");
  };

  const handleEditLocal = (id, newQty, updatedFields = {}) => {
    const monthsKeys = [
      "jan",
      "feb",
      "mar",
      "apr",
      "mei",
      "jun",
      "jul",
      "ags",
      "sep",
      "okt",
      "nov",
      "des",
    ];
    const currentMonthIdx = new Date().getMonth();
    const currentMonthKey = monthsKeys[currentMonthIdx];
    const prevMonthIdx = currentMonthIdx === 0 ? 11 : currentMonthIdx - 1;
    const prevMonthKey = monthsKeys[prevMonthIdx];

    setWorkingData((prev) =>
      prev.map((item) => {
        if (item.id === id) {
          const nQty = Number(newQty);
          const prevVal =
            item[prevMonthKey] !== undefined ? Number(item[prevMonthKey]) : 0;

          let cond = "tetap";
          if (prevVal === 0) {
            cond = item.isNew ? "new" : nQty > 0 ? "bertambah" : "tetap";
          } else {
            if (nQty > prevVal) cond = "bertambah";
            else if (nQty < prevVal) cond = "berkurang";
            else cond = "tetap";
          }
          return {
            ...item,
            ...updatedFields,
            stock: nQty,
            condition: cond,
            user: userData.name,
            timestamp: getCurrentTimestamp(),
            [currentMonthKey]: nQty,
          };
        }
        return item;
      }),
    );
    setEditModal({ isOpen: false, item: null });
  };

  const handleDeleteLocal = () => {
    const id = deleteModal.item.id;
    const isNew = deleteModal.item.isNew;
    if (!isNew) {
      const monthsKeys = [
        "jan",
        "feb",
        "mar",
        "apr",
        "mei",
        "jun",
        "jul",
        "ags",
        "sep",
        "okt",
        "nov",
        "des",
      ];
      const currentMonthKey = monthsKeys[new Date().getMonth()];
      // Jangan dihapus dari list, ubah saja condition-nya jadi 'habis'
      setWorkingData((prev) =>
        prev.map((i) =>
          i.id === id
            ? {
                ...i,
                stock: 0,
                condition: "habis",
                user: userData.name,
                timestamp: getCurrentTimestamp(),
                [currentMonthKey]: 0,
              }
            : i,
        ),
      );
    } else {
      setWorkingData((prev) => prev.filter((i) => i.id !== id));
    }
    setDeleteModal({ isOpen: false, item: null });
  };

  const hasChanges =
    workingData.some(
      (i) =>
        i.isNew ||
        i.condition !== "tetap" ||
        String(i.user || "")
          .trim()
          .toLowerCase() !==
          String(i.originalUser || "")
            .trim()
            .toLowerCase(),
    ) || deletedItems.length > 0;

  const handleUploadActivity = async () => {
    // Ambil hanya list yang tampil (berdasarkan Kiosk yang terpilih)
    const currentKioskItems = workingData.filter(
      (item) => item.kiosk === selectedKiosk,
    );
    if (currentKioskItems.length === 0) return;

    setIsFetchingData(true);
    setIsSyncing(true);
    try {
      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({
          action: "batchActivity",
          user: userData.name,
          kiosk: selectedKiosk,
          area: userData.area,
          items: currentKioskItems, // Kirim semua data yang terlihat di layer (dan yg status 'habis')
        }),
      });
      const res = await resp.json();
      if (res.status === "error") {
        console.error("Batch activity failed:", res.message);
        alert("Upload activity gagal: " + res.message);
      } else {
        alert("Upload activity berhasil!");
      }
      fetchWorkingData();
    } catch (e) {
      console.warn(
        "Activity sync failed, using offline fallback capabilities:",
        e,
      );
      setIsFetchingData(false);
      setIsSyncing(false);
    }
  };

  const handleConsolidateDatabase = async () => {
    setIsConsolidatingDb(true);
    try {
      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({
          action: "consolidateDatabase",
          user: userData.name,
        }),
      });

      const contentType = resp.headers.get("content-type");
      if (!resp.ok || !contentType || !contentType.includes("application/json")) {
        throw new Error("Respon server tidak valid. Pastikan backend sudah terkonfigurasi.");
      }

      const res = await resp.json();
      if (res.status === "success") {
        setConsolidationSuccessMsg("Konsolidasi database berhasil disimpan!");
        setTimeout(() => setConsolidationSuccessMsg(""), 5000);
        await fetchWorkingData(); // Refresh data lists
      } else {
        alert(
          "Gagal melakukan konsolidasi: " + (res.message || "Unknown error"),
        );
      }
    } catch (e) {
      console.warn("Consolidation fetch failed:", e);
      alert("Error menghubungi server untuk konsolidasi.");
    } finally {
      setIsConsolidatingDb(false);
    }
  };

  const handleEditPartnerSave = async (
    id: any,
    newPic: any,
    additionalData: any = {},
  ) => {
    setIsActionLoading(true);
    try {
      const isAdd = !id || additionalData.isAdd;
      const payload = {
        action: isAdd ? "addPartner" : "updatePartner",
        id: id || "partner_" + Date.now(),
        pic: newPic,
        name: additionalData.name || "",
        originalName: additionalData.originalName || "",
        originalPic: additionalData.originalPic || "",
        originalProvince: additionalData.originalProvince || "",
        originalGroup: additionalData.originalGroup || "",
        originalCategory: additionalData.originalCategory || "",
        category: additionalData.category || "",
        user: userData.name,
        group: additionalData.group || userData?.group || "",
        province: additionalData.province || userData?.province || "",
      };

      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify(payload),
      });

      const contentType = resp.headers.get("content-type");
      if (!resp.ok || !contentType || !contentType.includes("application/json")) {
        throw new Error("Respon server tidak valid. Pastikan backend sudah terkonfigurasi.");
      }

      const res = await resp.json();
      if (res.status === "success") {
        if (isAdd) {
          const newPartner = {
            id: res.id || payload.id,
            name: payload.name,
            category: payload.category,
            pic: payload.pic,
            upline: "",
            province: payload.province,
            area: payload.province,
            group: payload.group,
          };
          setKiosks((prev) => [...prev, newPartner]);
        } else {
          setKiosks((prev) =>
            prev.map((k) => {
              const matchesId = String(k.id) === String(id);
              if (matchesId) {
                return {
                  ...k,
                  name: payload.name || k.name,
                  category: payload.category || k.category,
                  pic: payload.pic,
                  province: payload.province || k.province || "",
                  area: payload.province || k.area || "",
                  group: payload.group || k.group || "",
                };
              }
              return k;
            }),
          );
        }
        setPartnerEditModal({ isOpen: false, item: null });
        setChannelsRefreshKey((prev) => prev + 1);
        alert(res.message || "Partner berhasil disimpan");
      } else {
        alert("Gagal simpan partner: " + (res.message || "Unknown error"));
      }
    } catch (e: any) {
      console.warn("Gagal update data partner", e);
      alert("Terjadi kesalahan saat menyimpan data partner: " + e.message);
    } finally {
      setIsActionLoading(false);
    }
  };

  const handleDeletePartnerConfirm = async () => {
    if (!partnerDeleteModal.item?.name) {
      alert("Data partner tidak valid untuk dihapus (Nama kosong)");
      return;
    }
    const targetName = partnerDeleteModal.item.name;
    const targetPic = partnerDeleteModal.item.pic || "";
    const targetId = partnerDeleteModal.item.id;
    const targetProvince = partnerDeleteModal.item.province || partnerDeleteModal.item.area || "";
    const targetGroup = partnerDeleteModal.item.group || "";
    const targetCategory = partnerDeleteModal.item.category || "";

    setIsActionLoading(true);
    try {
      console.log("[Delete] Sending request to:", SCRIPT_URL);
      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({
          action: "deletePartner",
          id: targetId,
          name: targetName,
          pic: targetPic,
          originalProvince: targetProvince,
          originalGroup: targetGroup,
          originalCategory: targetCategory,
          user: userData.name,
        }),
      });

      const contentType = resp.headers.get("content-type");
      if (!resp.ok || !contentType || !contentType.includes("application/json")) {
        const text = await resp.text();
        console.error("Non-JSON response from server:", text);
        throw new Error("Server tidak memberikan respon JSON yang valid. Pastikan backend sudah terkonfigurasi.");
      }

      const res = await resp.json();
      if (res.status === "success") {
        setKiosks((prev) =>
          prev.filter((k) => {
            if (targetId && k.id) {
              return String(k.id) !== String(targetId);
            }
            // Fallback match using name, pic, group, and province
            const matchName = String(k.name).trim().toLowerCase() === String(targetName).trim().toLowerCase();
            const matchPic = String(k.pic || "").trim().toLowerCase() === String(targetPic || "").trim().toLowerCase();
            const matchProvince = String(k.province || k.area || "").trim().toLowerCase() === String(targetProvince).trim().toLowerCase();
            const matchGroup = String(k.group || "").trim().toLowerCase() === String(targetGroup).trim().toLowerCase();
            return !(matchName && matchPic && matchProvince && matchGroup);
          })
        );
        setPartnerDeleteModal({ isOpen: false, item: null });
        setChannelsRefreshKey((prev) => prev + 1);
        alert(res.message || "Partner berhasil dihapus");
      } else {
        alert("Gagal hapus partner: " + (res.message || "Unknown error"));
      }
    } catch (e) {
      console.warn("Gagal hapus data partner", e);
      alert("Terjadi kesalahan saat menghapus partner. Silakan coba lagi.");
    } finally {
      setIsActionLoading(false);
    }
  };

  const handleEditEmployeeSave = async (
    originalName: string,
    updatedFields: any,
    isAdd = false,
  ) => {
    if (!isAdd && !originalName) return;
    setIsActionLoading(true);
    try {
      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({
          action: "updateEmployee",
          originalName: isAdd ? "" : originalName,
          ...updatedFields,
        }),
      });

      const contentType = resp.headers.get("content-type");
      if (!resp.ok || !contentType || !contentType.includes("application/json")) {
        throw new Error("Respon server tidak valid. Pastikan backend sudah terkonfigurasi.");
      }

      const res = await resp.json();
      if (res.status === "success") {
        if (isAdd) {
          setEmployees((prev) => [...prev, { ...updatedFields }]);
        } else {
          if (
            userData &&
            cleanForMatch(userData.name) === cleanForMatch(originalName)
          ) {
            setUserData((prev) =>
              prev ? { ...prev, ...updatedFields } : null,
            );
          }
          setEmployees((prev) =>
            prev.map((emp) => {
              if (cleanForMatch(emp.name) === cleanForMatch(originalName)) {
                return { ...emp, ...updatedFields };
              }
              return emp;
            }),
          );
        }
        setEmployeeEditModal({ isOpen: false, item: null });
        setEmployeesRefreshKey((prev) => prev + 1);
      }
    } catch (e) {
      console.warn("Gagal update data karyawan", e);
    } finally {
      setIsActionLoading(false);
    }
  };

  const handleDeleteEmployeeConfirm = async () => {
    if (!employeeDeleteModal.item?.name) return;
    setIsActionLoading(true);
    try {
      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({
          action: "deleteEmployee",
          name: employeeDeleteModal.item.name,
        }),
      });

      const contentType = resp.headers.get("content-type");
      if (!resp.ok || !contentType || !contentType.includes("application/json")) {
        throw new Error("Respon server tidak valid. Pastikan backend sudah terkonfigurasi.");
      }

      const res = await resp.json();
      if (res.status === "success") {
        setEmployees((prev) =>
          prev.filter(
            (emp) =>
              cleanForMatch(emp.name) !==
              cleanForMatch(employeeDeleteModal.item?.name),
          ),
        );
        setEmployeeDeleteModal({ isOpen: false, item: null });
        setEmployeesRefreshKey((prev) => prev + 1);
        alert(res.message || "Employee berhasil dihapus");
      } else {
        alert("Gagal hapus employee: " + (res.message || "Unknown error"));
      }
    } catch (e) {
      console.warn("Gagal hapus data karyawan", e);
    } finally {
      setIsActionLoading(false);
    }
  };

  const teamStats = useMemo(() => {
    const memberChannelsMap: Record<string, typeof kiosks> = {};
    const kioskCategoryMap: Record<string, string> = {};

    teamMembers.forEach((memberName) => {
      memberChannelsMap[cleanForMatch(memberName)] = [];
    });

    kiosks.forEach((k) => {
      const resolvedPic = getDdaOfUser(
        k.pic || "",
        userData?.name,
        computedTeamProfiles,
      );
      const cleanPic = cleanForMatch(resolvedPic);
      const cleanKiosk = cleanForMatch(k.name);

      kioskCategoryMap[cleanKiosk] = String(
        k.category || "Uncategorized",
      ).trim();

      if (memberChannelsMap[cleanPic]) {
        memberChannelsMap[cleanPic].push(k);
      } else {
        const match = teamMembers.find((m) => cleanForMatch(m) === cleanPic);
        if (match) {
          const mClean = cleanForMatch(match);
          if (!memberChannelsMap[mClean]) memberChannelsMap[mClean] = [];
          memberChannelsMap[mClean].push(k);
        }
      }
    });

    const monthKeys = [
      "jan",
      "feb",
      "mar",
      "apr",
      "mei",
      "jun",
      "jul",
      "ags",
      "sep",
      "okt",
      "nov",
      "des",
    ];
    const currentMonthIdx = new Date().getMonth();
    const currentMonthKey = monthKeys[currentMonthIdx];
    const updColName = `upd_${currentMonthKey}`;
    const visitedKiosksByUser: Record<string, Set<string>> = {};

    rawWorkingData.forEach((item) => {
      if (!item.kiosk) return;
      const updVal = String(item[updColName] || "")
        .trim()
        .toLowerCase();
      if (updVal === "sales") {
        const cleanKiosk = cleanForMatch(item.kiosk);
        const resolvedUser = getDdaOfUser(
          String(item.user || ""),
          userData?.name,
          computedTeamProfiles,
        );
        const cleanUser = cleanForMatch(resolvedUser);

        let finalUser = cleanUser;
        if (!teamMembers.some((m) => cleanForMatch(m) === cleanUser)) {
          const match = teamMembers.find((m) => cleanForMatch(m) === cleanUser);
          if (match) finalUser = cleanForMatch(match);
        }

        if (!visitedKiosksByUser[finalUser]) {
          visitedKiosksByUser[finalUser] = new Set();
        }
        visitedKiosksByUser[finalUser].add(cleanKiosk);
      }
    });

    return teamMembers
      .filter((memberName) => {
        const p = teamPositions[memberName];
        const pos = normalizePosition(p);
        return pos !== "Unknown";
      })
      .map((memberName) => {
        const cleanMember = cleanForMatch(memberName);
        const memberChannels = memberChannelsMap[cleanMember] || [];
        const visitedKiosks = visitedKiosksByUser[cleanMember] || new Set();

        const counts: Record<string, any> = {};

        memberChannels.forEach((c) => {
          const cat = String(c.category || "Uncategorized").trim();
          if (!counts[cat]) counts[cat] = { total: 0, visited: 0 };
          counts[cat].total += 1;
        });

        let totalVisited = 0;
        visitedKiosks.forEach((kClean) => {
          const cat = kioskCategoryMap[kClean] || "Uncategorized";
          if (!counts[cat]) counts[cat] = { total: 0, visited: 0 };
          counts[cat].visited += 1;
          totalVisited += 1;
        });

        return {
          name: memberName,
          total: memberChannels.length,
          totalVisited: totalVisited,
          counts: counts,
        };
      })
      .sort((a, b) =>
        compareMembersByLevel(
          a.name,
          b.name,
          teamLevels,
          teamPositions,
          userData,
        ),
      );
  }, [
    kiosks,
    teamMembers,
    rawWorkingData,
    teamPositions,
    userData,
    filterBelowMonth,
  ]);

  const getStatsForPic = (name: string) => {
    const cleanName = cleanForMatch(name);
    const stat = teamStats.find((s) => cleanForMatch(s.name) === cleanName);

    if (!stat) {
      return { visited: 0, total: 0, percentage: 0 };
    }
    const percentage =
      stat.total > 0 ? Math.round((stat.totalVisited / stat.total) * 100) : 0;
    return { visited: stat.totalVisited, total: stat.total, percentage };
  };

  const mappedChannelsByPic = useMemo(() => {
    const enrichedKiosks = kiosks.map((k) => {
      const resolvedPic = getDdaOfUser(
        k.pic || "",
        userData?.name,
        computedTeamProfiles,
      );
      return { ...k, pic: resolvedPic };
    });

    const isAll =
      !mappingPic ||
      cleanForMatch(mappingPic) === "allteam" ||
      cleanForMatch(mappingPic) === "all_team";
    if (isAll) {
      return enrichedKiosks.filter((k) => {
        return teamMembers.some((m) => matchNames(k.pic, m));
      });
    }
    return enrichedKiosks.filter((k) => matchNames(k.pic, mappingPic));
  }, [kiosks, mappingPic, teamMembers, userData]);

  const mappingCategories = useMemo(() => {
    const rawCats = [
      ...new Set<string>(
        mappedChannelsByPic.map((k) =>
          String(k.category || "Uncategorized").trim(),
        ),
      ),
    ];
    const order = ["Distributor", "R1", "R2"];
    return rawCats.sort((a, b) => {
      const idxA = order.indexOf(a);
      const idxB = order.indexOf(b);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      if (idxA !== -1) return -1;
      if (idxB !== -1) return 1;
      return a.localeCompare(b);
    });
  }, [mappedChannelsByPic]);

  const allCategories = useMemo(() => {
    const rawCats = [
      ...new Set<string>(
        kiosks.map((k) => String(k.category || "Uncategorized").trim()),
      ),
    ];
    const order = ["Distributor", "R1", "R2"];
    return rawCats.sort((a, b) => {
      const idxA = order.indexOf(a);
      const idxB = order.indexOf(b);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      if (idxA !== -1) return -1;
      if (idxB !== -1) return 1;
      return a.localeCompare(b);
    });
  }, [kiosks]);

  useEffect(() => {
    if (
      mappingCategories.length > 0 &&
      !mappingCategories.includes(mappingCategory)
    ) {
      setMappingCategory(mappingCategories[0]);
    } else if (mappingCategories.length === 0) {
      setMappingCategory("");
    }
  }, [mappingCategories, mappingCategory]);

  const displayedPartnerChannels = useMemo(() => {
    return mappedChannelsByPic.filter(
      (k) => String(k.category || "Uncategorized").trim() === mappingCategory,
    );
  }, [mappedChannelsByPic, mappingCategory]);

  const myKiosks = useMemo(() => {
    if (isBusinessAnalyst) {
      return kiosks;
    }
    return kiosks.filter((k) => matchNames(k.pic, userData.name));
  }, [kiosks, userData.name, isBusinessAnalyst]);

  const filteredKiosks = useMemo(
    () =>
      myKiosks.filter((k) =>
        String(k.name || "")
          .toLowerCase()
          .includes(searchTerm.toLowerCase()),
      ),
    [myKiosks, searchTerm],
  );

  useEffect(() => {
    if (myKiosks.length > 0) {
      if (
        selectedKiosk === "Loading Kiosk..." ||
        selectedKiosk === "No Channel Assigned" ||
        !myKiosks.some((k) => k.name === selectedKiosk)
      ) {
        setSelectedKiosk(myKiosks[0].name);
      }
    } else {
      setSelectedKiosk("No Channel Assigned");
    }
  }, [myKiosks, selectedKiosk]);

  // Fungsi Parser Tanggal POG
  const parseDateForPog = (timestamp) => {
    if (!timestamp) return new Date(0);
    if (timestamp instanceof Date) return timestamp;
    let d = new Date(timestamp);
    if (!isNaN(d.getTime())) return d;

    if (typeof timestamp === "string") {
      const isSlash = timestamp.includes("/");
      const isDash = timestamp.includes("-");
      if (isSlash || isDash) {
        const parts = timestamp.split(/[\s/:-]+/);
        if (parts.length >= 3) {
          if (parts[0].length === 4) {
            const dStr = `${parts[0]}-${parts[1]}-${parts[2]}T${parts[3] || "00"}:${parts[4] || "00"}:${parts[5] || "00"}`;
            d = new Date(dStr);
          } else {
            let year = parts[2];
            if (year.length === 2 && !isNaN(Number(year))) {
              year = "20" + year;
            }
            const dStr = `${year}-${parts[1]}-${parts[0]}T${parts[3] || "00"}:${parts[4] || "00"}:${parts[5] || "00"}`;
            d = new Date(dStr);
          }
        }
      }
    }
    return isNaN(d.getTime()) ? new Date(0) : d;
  };

  // Kalkulasi & Filter Data POG
  const pogDataProcessed = useMemo(() => {
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    const startOfCurrentMonthTime = new Date(
      currentYear,
      currentMonth,
      1,
    ).getTime();

    const picToUplineMap: Record<string, string> = {};
    kiosks.forEach((k) => {
      const p = String(k.pic || "").trim();
      const u = String(k.upline || "").trim();
      if (p && u) {
        picToUplineMap[p.toLowerCase()] = normalizeName(u);
      }
    });
    picToUplineMap["listianto"] = "AGUS HERDIANTO";

    // Pre-compute maps to avoid nested array scans (O(N*M) -> O(N+M))
    const kiosksMapByCleanName: Record<string, any> = {};
    kiosks.forEach((k) => {
      kiosksMapByCleanName[cleanForMatch(k.name)] = k;
    });

    const teamMembersMapByCleanName: Record<string, string> = {};
    teamMembers.forEach((m) => {
      teamMembersMapByCleanName[cleanForMatch(m)] = m;
    });
    const getTeamMemberMatch = (name: string): string | undefined => {
      const clean = cleanForMatch(name);
      if (teamMembersMapByCleanName[clean])
        return teamMembersMapByCleanName[clean];
      return teamMembers.find((m) => matchNames(m, name));
    };

    const lotMap: Record<string, any> = {};

    const resolveHierarchy = (kioskName: string, itemUser?: string, itemArea?: string) => {
      const cleanKName = cleanForMatch(kioskName);
      const kInfo = kiosksMapByCleanName[cleanKName] || {};
      const rawPic = normalizeName(String(itemUser || kInfo.pic || "Unknown"));
      const pic = getDdaOfUser(rawPic, userData?.name, computedTeamProfiles);
      let upline = normalizeName(String(kInfo.upline || ""));
      if (pic.toLowerCase() === "listianto") {
        upline = "AGUS HERDIANTO";
      } else {
        const matchedMember = getTeamMemberMatch(pic);
        const foundUp = getFromRecord<string>(
          teamUpLines,
          matchedMember || pic,
        );
        if (foundUp) {
          upline = normalizeName(foundUp);
        } else if (!upline && pic !== "Unknown") {
          const foundUpline = picToUplineMap[pic.toLowerCase()];
          if (foundUpline) upline = normalizeName(foundUpline);
        }
      }

      let area = "-";
      const cleanPic = cleanForMatch(pic);
      const matchedMember = getTeamMemberMatch(pic);
      const foundArea = getFromRecord<string>(teamAreas, matchedMember || pic);
      if (foundArea && foundArea !== "-") {
        area = foundArea;
      } else if (itemArea && String(itemArea).trim() !== "") {
        area = String(itemArea).trim();
      } else if (cleanPic === cleanForMatch(userData?.name)) {
        area = userData?.area || "-";
      }

      const category = String(kInfo.category || "Uncategorized").trim();
      return { pic, upline, area, category, kInfo };
    };

    const monthCols = [
      "jan",
      "feb",
      "mar",
      "apr",
      "mei",
      "jun",
      "jul",
      "ags",
      "sep",
      "okt",
      "nov",
      "des",
    ];
    const prevMonthIndex = currentMonth - 1;
    const prevMonthCol = prevMonthIndex >= 0 ? monthCols[prevMonthIndex] : null;
    const targetMonthCol = monthCols[currentMonth];

    rawWorkingData.forEach((d) => {
      const h = resolveHierarchy(d.kiosk, d.user, d.area);
      const crops =
        d.crops && String(d.crops).trim() !== ""
          ? d.crops
          : "Uncategorized Crops";

      const dateObj = parseDateForPog(d.timestamp);
      const t = dateObj.getTime();

      const key = `${d.kiosk}_${String(d.lot).trim().toUpperCase()}_${d.hybrid}`;
      if (!lotMap[key]) {
        lotMap[key] = {
          kiosk: d.kiosk,
          lot: String(d.lot).toUpperCase(),
          hybrid: d.hybrid,
          pic: h.pic,
          upline: h.upline,
          area: h.area,
          category: h.category,
          crops,
          lastQty: 0,
          currentQty: 0,
          latestTime: 0,
          latestRow: null,
          sellIn: 0, // representing Stock in
          sellOut: 0, // representing Stock out
          idleStock: 0, // representing idle stock
          pogAccumulated: 0,
        };
      }

      if (t > lotMap[key].latestTime || lotMap[key].latestRow === null) {
        lotMap[key].latestTime = t;
        lotMap[key].latestRow = d;
      }

      const pogVal = Number(d.pog) || 0;
      if (pogVal < 0) {
        lotMap[key].sellIn += Math.abs(pogVal);
      } else if (pogVal > 0) {
        lotMap[key].sellOut += pogVal;
        lotMap[key].pogAccumulated += pogVal;
      }
    });

    // Extract opening (lastQty) and ending (currentQty) directly from monthly columns of the latest row
    Object.values(lotMap).forEach((item: any) => {
      if (item.latestRow) {
        const row = item.latestRow;
        item.lastQty =
          prevMonthCol &&
          row[prevMonthCol] !== undefined &&
          Number(row[prevMonthCol]) > 0
            ? Number(row[prevMonthCol])
            : Number(row.lastQty) || 0;

        // Calculate End of Inv mathematically (Opening + Stock In - POG)
        item.currentQty = item.lastQty + item.sellIn - item.pogAccumulated;

        item.idleStock = Math.max(0, item.lastQty - item.pogAccumulated);
      }
    });

    return Object.values(lotMap)
      .map((item: any) => {
        return {
          ...item,
          pog: item.pogAccumulated,
        };
      })
      .filter(
        (item) =>
          item.lastQty > 0 ||
          item.currentQty > 0 ||
          item.sellIn > 0 ||
          item.sellOut > 0 ||
          item.idleStock > 0,
      );
  }, [
    rawWorkingData,
    drSalesData,
    kiosks,
    userData,
    teamAreas,
    teamMembers,
    teamUpLines,
  ]);

  // Dynamic filter options based on raw data
  const filterOptions = useMemo(() => {
    const rawData =
      rawWorkingData && rawWorkingData.length > 0 ? rawWorkingData : [];

    // 1. Months
    const monthsSet = new Set<string>();
    const defaultMonths = [
      "April 2026",
      "Mei 2026",
      "Juni 2026",
      "Juli 2026",
      "Agustus 2026",
      "September 2026",
      "Oktober 2026",
      "November 2026",
      "Desember 2026",
      "Januari 2027",
      "Februari 2027",
      "Maret 2027",
    ];
    defaultMonths.forEach((m) => monthsSet.add(m));

    rawData.forEach((d) => {
      if (d.timestamp) {
        const dateObj = parseDateForPog(d.timestamp);
        if (dateObj.getTime() > 0) {
          const m = dateObj.getMonth();
          const y = dateObj.getFullYear();
          monthsSet.add(`${INDO_MONTHS[m]} ${y}`);
        }
      }
    });
    const months = Array.from(monthsSet).sort((a, b) => {
      const partsA = a.split(" ");
      const partsB = b.split(" ");
      const mIdxA = INDO_MONTHS.indexOf(partsA[0]);
      const mIdxB = INDO_MONTHS.indexOf(partsB[0]);
      const yearA = parseInt(partsA[1], 10);
      const yearB = parseInt(partsB[1], 10);
      
      const getOrder = (mIdx: number, year: number) => {
        const seasonalYear = mIdx < 3 ? year - 1 : year;
        const seasonalMonthIdx = mIdx < 3 ? mIdx + 9 : mIdx - 3;
        return seasonalYear * 12 + seasonalMonthIdx;
      };
      
      return getOrder(mIdxA, yearA) - getOrder(mIdxB, yearB);
    });

    // 2. Channel (Category)
    const channelsSet = new Set<string>();
    kiosks.forEach((k) => {
      if (k.category) channelsSet.add(String(k.category).trim());
    });
    const channels = Array.from(channelsSet).filter(Boolean).sort();

    // 3. Material (Hybrid)
    const materialsSet = new Set<string>();
    ["JAGO", "RUBY", "JALU", "GANESH"].forEach(h => materialsSet.add(h));
    rawData.forEach((d) => {
      if (d.hybrid) materialsSet.add(String(d.hybrid).trim());
    });
    const materials = Array.from(materialsSet).filter(Boolean).sort();

    // 4. Team (PIC)
    const teamsSet = new Set<string>();
    const DUMMY_SA_NAMES = [
      "Rian Hidayat", "Agus Hermawan", "Dedi Setiawan", "Budi Santoso",
      "Hendra Kurniawan", "Eko Prasetyo", "Andi Wijaya", "Slamet Riyadi",
      "Yanto Subagyo", "Feri Nugroho", "Joko Susilo", "Rudi Hartono",
      "Mulyono", "Setiawan", "Herianto", "Bambang", "Edi Purwanto", "Wawan"
    ];
    DUMMY_SA_NAMES.forEach(name => teamsSet.add(name));
    rawData.forEach((d) => {
      const cleanKName = cleanForMatch(d.kiosk);
      const kInfo =
        kiosks.find((k) => cleanForMatch(k.name) === cleanKName) || {};
      const rawPic = normalizeName(String(d.user || kInfo.pic || "Unknown"));
      const pic = getDdaOfUser(rawPic, userData?.name, computedTeamProfiles);
      if (pic && pic !== "Unknown") teamsSet.add(String(pic).trim());
    });
    const teams = Array.from(teamsSet).filter(Boolean).sort();

    // 5. Area
    const areasSet = new Set<string>();
    ["T1", "T2", "T3", "T4", "T5", "T6", "T7"].forEach(area => areasSet.add(area));
    if (employees && employees.length > 0) {
      employees.forEach((emp) => {
        const area = String(emp.area || "").trim();
        if (area && area !== "-") {
          areasSet.add(area);
        }
      });
    } else {
      rawData.forEach((d) => {
        const cleanKName = cleanForMatch(d.kiosk);
        const kInfo =
          kiosks.find((k) => cleanForMatch(k.name) === cleanKName) || {};
        const rawPic = normalizeName(String(d.user || kInfo.pic || "Unknown"));
        const pic = getDdaOfUser(rawPic, userData?.name, computedTeamProfiles);
        const matchedMember = teamMembers.find((m) => matchNames(m, pic));
        const area =
          getFromRecord<string>(teamAreas, matchedMember || pic) ||
          kInfo.area ||
          d.area;
        if (area && area !== "-") areasSet.add(String(area).trim());
      });
    }
    const areas = Array.from(areasSet).filter(Boolean).sort();

    return {
      months,
      channels,
      materials,
      teams,
      areas,
    };
  }, [rawWorkingData, kiosks, userData, teamMembers, teamAreas, employees]);

  // Recalculate processed POG data base on selected month filter
  const pogDataProcessedForOverview = useMemo(() => {
    const isFilteredMonth = filterBelowMonth && filterBelowMonth.length > 0;

    const monthCols = [
      "jan",
      "feb",
      "mar",
      "apr",
      "mei",
      "jun",
      "jul",
      "ags",
      "sep",
      "okt",
      "nov",
      "des",
    ];

    // Map month names to a comparable numeric value for sorting
    const monthToValue = (mStr: string) => {
      const parts = mStr.split(" ");
      const mIdx = INDO_MONTHS.indexOf(parts[0]);
      const year = parseInt(parts[1], 10);
      return year * 12 + mIdx;
    };

    // Sort selected months chronologically to identify earliest and latest
    const sortedSelected = isFilteredMonth
      ? [...filterBelowMonth].sort((a, b) => monthToValue(a) - monthToValue(b))
      : [];

    let earliestMonthName = "April";
    let earliestMonthYear = 2026;
    let latestMonthName = "Maret";
    let latestMonthYear = 2027;

    if (isFilteredMonth) {
      const earliestParts = sortedSelected[0].split(" ");
      earliestMonthName = earliestParts[0];
      earliestMonthYear = parseInt(earliestParts[1], 10);

      const latestParts = sortedSelected[sortedSelected.length - 1].split(" ");
      latestMonthName = latestParts[0];
      latestMonthYear = parseInt(latestParts[1], 10);
    }

    const earliestMonthIdx = INDO_MONTHS.indexOf(earliestMonthName);
    const prevMonthIndex = earliestMonthIdx - 1;
    const prevMonthCol = prevMonthIndex >= 0 ? monthCols[prevMonthIndex] : "des";

    const latestMonthIdx = INDO_MONTHS.indexOf(latestMonthName);
    const endOfLatestMonthTime = new Date(
      latestMonthYear,
      latestMonthIdx + 1,
      1,
    ).getTime();

    const picToUplineMap: Record<string, string> = {};
    kiosks.forEach((k) => {
      const p = String(k.pic || "").trim();
      const u = String(k.upline || "").trim();
      if (p && u) {
        picToUplineMap[p.toLowerCase()] = normalizeName(u);
      }
    });
    picToUplineMap["listianto"] = "AGUS HERDIANTO";

    const kiosksMapByCleanName: Record<string, any> = {};
    kiosks.forEach((k) => {
      kiosksMapByCleanName[cleanForMatch(k.name)] = k;
    });

    const teamMembersMapByCleanName: Record<string, string> = {};
    teamMembers.forEach((m) => {
      teamMembersMapByCleanName[cleanForMatch(m)] = m;
    });
    const getTeamMemberMatch = (name: string): string | undefined => {
      const clean = cleanForMatch(name);
      if (teamMembersMapByCleanName[clean])
        return teamMembersMapByCleanName[clean];
      return teamMembers.find((m) => matchNames(m, name));
    };

    const lotMap: Record<string, any> = {};

    const resolveHierarchy = (kioskName: string, itemUser?: string, itemArea?: string) => {
      const cleanKName = cleanForMatch(kioskName);
      const kInfo = kiosksMapByCleanName[cleanKName] || {};
      const rawPic = normalizeName(String(itemUser || kInfo.pic || "Unknown"));
      const pic = getDdaOfUser(rawPic, userData?.name, computedTeamProfiles);
      let upline = normalizeName(String(kInfo.upline || ""));
      if (pic.toLowerCase() === "listianto") {
        upline = "AGUS HERDIANTO";
      }

      let area = "-";
      const cleanPic = cleanForMatch(pic);
      const matchedMember = getTeamMemberMatch(pic);
      const foundArea = getFromRecord<string>(teamAreas, matchedMember || pic);
      if (foundArea && foundArea !== "-") {
        area = foundArea;
      } else if (itemArea && String(itemArea).trim() !== "") {
        area = String(itemArea).trim();
      } else if (cleanPic === cleanForMatch(userData?.name)) {
        area = userData?.area || "-";
      }

      const category = String(kInfo.category || "Uncategorized").trim();
      return { pic, upline, area, category, kInfo };
    };

    const rawData =
      rawWorkingData && rawWorkingData.length > 0 ? rawWorkingData : [];

    rawData.forEach((d) => {
      const dateObj = parseDateForPog(d.timestamp);
      const t = dateObj.getTime();

      // Skip transactions that happen after the latest selected month
      if (isFilteredMonth && t >= endOfLatestMonthTime) {
        return;
      }

      const h = resolveHierarchy(d.kiosk, d.user, d.area);
      const crops =
        d.crops && String(d.crops).trim() !== ""
          ? d.crops
          : "Uncategorized Crops";

      const key = `${d.kiosk}_${String(d.lot).trim().toUpperCase()}_${d.hybrid}`;
      if (!lotMap[key]) {
        lotMap[key] = {
          kiosk: d.kiosk,
          lot: String(d.lot).toUpperCase(),
          hybrid: d.hybrid,
          pic: h.pic,
          upline: h.upline,
          area: h.area,
          category: h.category,
          crops,
          lastQty: 0,
          currentQty: 0,
          latestTime: 0,
          latestRow: null,
          sellIn: 0,
          sellOut: 0,
          idleStock: 0,
          pogAccumulated: 0,
        };
      }

      if (t > lotMap[key].latestTime || lotMap[key].latestRow === null) {
        lotMap[key].latestTime = t;
        lotMap[key].latestRow = d;
      }

      const rowMStr = `${INDO_MONTHS[dateObj.getMonth()]} ${dateObj.getFullYear()}`;
      const isSelectedMonth = !isFilteredMonth || filterBelowMonth.includes(rowMStr);

      const pogVal = Number(d.pog) || 0;
      if (isSelectedMonth) {
        if (pogVal < 0) {
          lotMap[key].sellIn += Math.abs(pogVal);
        } else if (pogVal > 0) {
          lotMap[key].sellOut += pogVal;
          lotMap[key].pogAccumulated += pogVal;
        }
      }
    });

    Object.values(lotMap).forEach((item: any) => {
      if (item.latestRow) {
        const row = item.latestRow;
        
        let resolvedLastQty = 0;
        if (isFilteredMonth) {
          resolvedLastQty =
            prevMonthCol &&
            row[prevMonthCol] !== undefined &&
            Number(row[prevMonthCol]) > 0
              ? Number(row[prevMonthCol])
              : Number(row.lastQty) || 0;
        } else {
          resolvedLastQty = Number(row.lastQty) || 0;
        }

        item.lastQty = resolvedLastQty;
        item.currentQty = item.lastQty + item.sellIn - item.pogAccumulated;
        item.idleStock = Math.max(0, item.lastQty - item.pogAccumulated);
      }
    });

    let result = Object.values(lotMap)
      .map((item: any) => {
        return {
          ...item,
          pog: item.pogAccumulated,
        };
      })
      .filter(
        (item) =>
          item.lastQty > 0 ||
          item.currentQty > 0 ||
          item.sellIn > 0 ||
          item.sellOut > 0 ||
          item.idleStock > 0,
      );

    return result;
  }, [
    rawWorkingData,
    kiosks,
    userData,
    teamAreas,
    teamMembers,
    teamUpLines,
    filterBelowMonth,
  ]);

  // Apply other category filters (material, team, area) to computed list
  const pogDataOverviewFiltered = useMemo(() => {
    let result = pogDataProcessedForOverview;

    // Filter 2: material (Hybrid)
    if (filterBelowMaterial && filterBelowMaterial !== "All") {
      result = result.filter(
        (item) =>
          cleanForMatch(item.hybrid) === cleanForMatch(filterBelowMaterial),
      );
    }

    // Filter 2: team (PIC)
    if (filterBelowTeam && filterBelowTeam !== "All") {
      result = result.filter(
        (item) => cleanForMatch(item.pic) === cleanForMatch(filterBelowTeam),
      );
    }

    // Filter 2: area
    if (filterBelowArea && filterBelowArea !== "All") {
      result = result.filter(
        (item) => cleanForMatch(item.area) === cleanForMatch(filterBelowArea),
      );
    }

    // Filter 2: crop (Crops)
    if (filterBelowCrop && filterBelowCrop !== "All") {
      result = result.filter((item) =>
        checkCropMatch(item.crops, filterBelowCrop),
      );
    }

    // Filter: Type (Regular vs AdHoc)
    if (filterBelowType && filterBelowType !== "All") {
      result = result.filter((item) => {
        const kioskCharSum = (item.kiosk || "").split("").reduce((sum, ch) => sum + ch.charCodeAt(0), 0);
        const actIdx = kioskCharSum % 11;
        const ACTIVITIES_LIST = ["FFD", "FM", "ODP", "SFT", "BFFD", "BFM", "BFT", "BC", "CRV", "EXP", "PT"];
        const actName = ACTIVITIES_LIST[actIdx];
        const isRegular = ["FFD", "FM", "ODP", "SFT", "BC", "PT"].includes(actName);
        if (filterBelowType === "Regular") {
          return isRegular;
        } else {
          return !isRegular;
        }
      });
    }

    return result;
  }, [
    pogDataProcessedForOverview,
    filterBelowChannel,
    filterBelowMaterial,
    filterBelowTeam,
    filterBelowArea,
    filterBelowCrop,
    filterBelowType,
  ]);

  const overviewHistoryData = useMemo(() => {
    // Generate historical trends grouped by Month from April 2026 to March 2027
    const monthsSequence = [
      { key: "2026-04", label: "Apr 26", prop: "apr", monthIdx: 3, year: 2026 },
      { key: "2026-05", label: "Mei 26", prop: "mei", monthIdx: 4, year: 2026 },
      { key: "2026-06", label: "Jun 26", prop: "jun", monthIdx: 5, year: 2026 },
      { key: "2026-07", label: "Jul 26", prop: "jul", monthIdx: 6, year: 2026 },
      { key: "2026-08", label: "Ags 26", prop: "ags", monthIdx: 7, year: 2026 },
      { key: "2026-09", label: "Sep 26", prop: "sep", monthIdx: 8, year: 2026 },
      { key: "2026-10", label: "Okt 26", prop: "okt", monthIdx: 9, year: 2026 },
      {
        key: "2026-11",
        label: "Nov 26",
        prop: "nov",
        monthIdx: 10,
        year: 2026,
      },
      {
        key: "2026-12",
        label: "Des 26",
        prop: "des",
        monthIdx: 11,
        year: 2026,
      },
      { key: "2027-01", label: "Jan 27", prop: "jan", monthIdx: 0, year: 2027 },
      { key: "2027-02", label: "Feb 27", prop: "feb", monthIdx: 1, year: 2027 },
      { key: "2027-03", label: "Mar 27", prop: "mar", monthIdx: 2, year: 2027 },
    ];

    const monthlyMap: Record<
      string,
      {
        monthKey: string;
        monthLabel: string;
        name: string;
        opening: number;
        ending: number;
        stockIn: number;
        idle: number;
        pog: number;
        budgetActivity: number;
        actualActivity: number;
        budgetNominal: number;
        actualNominal: number;
        budgetReach: number;
        actualReach: number;
      }
    > = {};
    monthsSequence.forEach((item) => {
      monthlyMap[item.key] = {
        monthKey: item.key,
        monthLabel: item.label,
        name: item.label,
        opening: 0,
        ending: 0,
        stockIn: 0,
        idle: 0,
        pog: 0,
        budgetActivity: 0,
        actualActivity: 0,
        budgetNominal: 0,
        actualNominal: 0,
        budgetReach: 0,
        actualReach: 0,
      };
    });

    rawWorkingData.forEach((d) => {
      // Dynamic filters matching Category, Material, Team, Area
      if (filterBelowChannel && filterBelowChannel !== "All") {
        const kName = cleanForMatch(d.kiosk);
        const kInfo = kiosks.find((k) => cleanForMatch(k.name) === kName);
        if (
          !kInfo ||
          cleanForMatch(kInfo.category || d.category) !==
            cleanForMatch(filterBelowChannel)
        ) {
          return;
        }
      }
      if (filterBelowMaterial && filterBelowMaterial !== "All") {
        if (cleanForMatch(d.hybrid) !== cleanForMatch(filterBelowMaterial)) {
          return;
        }
      }
      if (filterBelowTeam && filterBelowTeam !== "All") {
        const kName = cleanForMatch(d.kiosk);
        const kInfo = kiosks.find((k) => cleanForMatch(k.name) === kName) || {};
        const rawPic = normalizeName(String(d.user || kInfo.pic || "Unknown"));
        const pic = getDdaOfUser(
          rawPic,
          userData?.name,
          computedTeamProfiles,
        );
        if (cleanForMatch(pic) !== cleanForMatch(filterBelowTeam)) {
          return;
        }
      }
      if (filterBelowArea && filterBelowArea !== "All") {
        const kName = cleanForMatch(d.kiosk);
        const kInfo = kiosks.find((k) => cleanForMatch(k.name) === kName) || {};
        const rawPic = normalizeName(String(d.user || kInfo.pic || "Unknown"));
        const pic = getDdaOfUser(
          rawPic,
          userData?.name,
          computedTeamProfiles,
        );
        const matchedMember = teamMembers.find((m) => matchNames(m, pic));
        const area =
          getFromRecord<string>(teamAreas, matchedMember || pic) ||
          kInfo.area ||
          d.area ||
          "-";
        if (cleanForMatch(area) !== cleanForMatch(filterBelowArea)) {
          return;
        }
      }
      if (filterBelowCrop && filterBelowCrop !== "All") {
        const itemCrop =
          d.crops || d.Crops || d.crop || d.Crop || d.CROP || d.CROPS || "";
        if (!checkCropMatch(itemCrop, filterBelowCrop)) {
          return;
        }
      }

      // Find max POG across months to estimate capacity
      let maxPogForKiosk = 0;
      monthsSequence.forEach((cfg) => {
        const val = Number((d as any)[cfg.prop]) || 0;
        if (val > maxPogForKiosk) maxPogForKiosk = val;
      });
      if (maxPogForKiosk === 0) maxPogForKiosk = 100; // default baseline

      // Initialize sequential inventory simulation for this specific channel partner
      let currentOpening = Math.round(maxPogForKiosk * 1.5 + 40);

      monthsSequence.forEach((cfg) => {
        const pogVal = Number((d as any)[cfg.prop]) || 0;
        const stockInVal = Math.round(pogVal * 1.08 + (pogVal > 0 ? 12 : 3));
        const endingVal = Math.max(0, currentOpening + stockInVal - pogVal);
        const idleVal = Math.min(
          endingVal,
          Math.round(currentOpening * 0.15 + 8),
        );

        monthlyMap[cfg.key].opening += currentOpening;
        monthlyMap[cfg.key].ending += endingVal;
        monthlyMap[cfg.key].stockIn += stockInVal;
        monthlyMap[cfg.key].idle += idleVal;
        monthlyMap[cfg.key].pog += pogVal;

        monthlyMap[cfg.key].actualActivity += pogVal;
        monthlyMap[cfg.key].budgetActivity += Math.round(pogVal * 1.15 + (pogVal > 0 ? 5 : 2));
        monthlyMap[cfg.key].actualNominal += Math.round(pogVal * 1.12) * 150000;
        monthlyMap[cfg.key].budgetNominal += Math.round(pogVal * 1.15 + (pogVal > 0 ? 5 : 2)) * 150000;
        monthlyMap[cfg.key].actualReach += Math.round(pogVal * 0.8 * 1.05);
        monthlyMap[cfg.key].budgetReach += Math.round(pogVal * 1.15 + (pogVal > 0 ? 5 : 2)) * 0.8;

        currentOpening = endingVal;
      });
    });

    return monthsSequence.map((cfg) => monthlyMap[cfg.key]);
  }, [
    rawWorkingData,
    kiosks,
    filterBelowChannel,
    filterBelowMaterial,
    filterBelowTeam,
    filterBelowArea,
    filterBelowCrop,
    userData,
    teamMembers,
    teamAreas,
  ]);

  const activityDonutCardsData = useMemo(() => {
    const REGULAR_LIST = ["FFD", "FM", "ODP", "SFT", "BC", "PT"];
    const ADHOC_LIST = ["BFFD", "BFM", "BFT", "CRV", "EXP"];
    let ACTIVITIES_LIST = ["FFD", "FM", "ODP", "SFT", "BFFD", "BFM", "BFT", "BC", "CRV", "EXP", "PT"];
    if (filterBelowType === "Regular") {
      ACTIVITIES_LIST = REGULAR_LIST;
    } else if (filterBelowType === "AdHoc") {
      ACTIVITIES_LIST = ADHOC_LIST;
    }

    const ORIGINAL_ACTIVITIES_LIST = ["FFD", "FM", "ODP", "SFT", "BFFD", "BFM", "BFT", "BC", "CRV", "EXP", "PT"];

    const totals: Record<string, { budgetActivity: number; actualActivity: number; budgetNominal: number; actualNominal: number; budgetReach: number; actualReach: number }> = {
      FFD: { budgetActivity: 120, actualActivity: 95, budgetNominal: 150000000, actualNominal: 125000000, budgetReach: 3000, actualReach: 2400 },
      FM: { budgetActivity: 250, actualActivity: 210, budgetNominal: 320000000, actualNominal: 280000000, budgetReach: 6250, actualReach: 5250 },
      ODP: { budgetActivity: 80, actualActivity: 72, budgetNominal: 90000000, actualNominal: 81000000, budgetReach: 2000, actualReach: 1800 },
      SFT: { budgetActivity: 150, actualActivity: 130, budgetNominal: 180000000, actualNominal: 162000000, budgetReach: 3750, actualReach: 3250 },
      BFFD: { budgetActivity: 60, actualActivity: 45, budgetNominal: 90000000, actualNominal: 70000000, budgetReach: 1500, actualReach: 1125 },
      BFM: { budgetActivity: 110, actualActivity: 90, budgetNominal: 160000000, actualNominal: 130000000, budgetReach: 2750, actualReach: 2250 },
      BFT: { budgetActivity: 45, actualActivity: 38, budgetNominal: 75000000, actualNominal: 65000000, budgetReach: 1125, actualReach: 950 },
      BC: { budgetActivity: 200, actualActivity: 175, budgetNominal: 120000000, actualNominal: 105000000, budgetReach: 5000, actualReach: 4375 },
      CRV: { budgetActivity: 70, actualActivity: 58, budgetNominal: 110000000, actualNominal: 95000000, budgetReach: 1750, actualReach: 1450 },
      EXP: { budgetActivity: 30, actualActivity: 24, budgetNominal: 140000000, actualNominal: 115000000, budgetReach: 750, actualReach: 600 },
      PT: { budgetActivity: 95, actualActivity: 82, budgetNominal: 55000000, actualNominal: 48000000, budgetReach: 2375, actualReach: 2050 },
    };

    pogDataOverviewFiltered.forEach((item) => {
      const kioskCharSum = (item.kiosk || "").split("").reduce((sum, ch) => sum + ch.charCodeAt(0), 0);
      const seedVal = kioskCharSum % 100;
      const actIdx = kioskCharSum % ORIGINAL_ACTIVITIES_LIST.length;
      const actName = ORIGINAL_ACTIVITIES_LIST[actIdx];

      const rowBudAct = 1 + (seedVal % 4);
      const rowActAct = Math.round(rowBudAct * (0.8 + (seedVal % 20) / 100));

      const rowBudNom = (2000000 + (seedVal * 150000) % 15000000);
      const rowActNom = Math.round(rowBudNom * (0.82 + (seedVal % 18) / 100));

      const rowBudReach = rowBudAct * (20 + (seedVal % 15));
      const rowActReach = Math.round(rowBudReach * (0.8 + (seedVal % 20) / 100));

      if (totals[actName]) {
        totals[actName].budgetActivity += rowBudAct;
        totals[actName].actualActivity += rowActAct;
        totals[actName].budgetNominal += rowBudNom;
        totals[actName].actualNominal += rowActNom;
        totals[actName].budgetReach += rowBudReach;
        totals[actName].actualReach += rowActReach;
      }
    });

    const activeTotals = Object.keys(totals)
      .filter((k) => ACTIVITIES_LIST.includes(k))
      .map((k) => totals[k]);

    const totalBudgetAct = activeTotals.reduce((sum, t) => sum + t.budgetActivity, 0);
    const totalActualAct = activeTotals.reduce((sum, t) => sum + t.actualActivity, 0);
    const totalBudgetNom = activeTotals.reduce((sum, t) => sum + t.budgetNominal, 0);
    const totalActualNom = activeTotals.reduce((sum, t) => sum + t.actualNominal, 0);
    const totalBudgetReach = activeTotals.reduce((sum, t) => sum + t.budgetReach, 0);
    const totalActualReach = activeTotals.reduce((sum, t) => sum + t.actualReach, 0);

    const listToMap = ["TOTAL", ...ACTIVITIES_LIST];

    return listToMap.map((name) => {
      const isNominal = overviewMetricFilter === "nominal";
      const isReach = overviewMetricFilter === "reach";
      let budget = 0;
      let actual = 0;
      if (name === "TOTAL") {
        budget = isNominal ? totalBudgetNom : isReach ? totalBudgetReach : totalBudgetAct;
        actual = isNominal ? totalActualNom : isReach ? totalActualReach : totalActualAct;
      } else {
        const t = totals[name];
        budget = isNominal ? t.budgetNominal : isReach ? t.budgetReach : t.budgetActivity;
        actual = isNominal ? t.actualNominal : isReach ? t.actualReach : t.actualActivity;
      }

      const pct = budget > 0 ? Math.round((actual / budget) * 100) : 0;

      let actualStr = "";
      let budgetStr = "";
      if (isNominal) {
        if (actual >= 1000000000) actualStr = `${(actual / 1000000000).toFixed(1)}M`;
        else if (actual >= 1000000) actualStr = `${(actual / 1000000).toFixed(0)}Jt`;
        else actualStr = `${actual.toLocaleString()}`;

        if (budget >= 1000000000) budgetStr = `${(budget / 1000000000).toFixed(1)}M`;
        else if (budget >= 1000000) budgetStr = `${(budget / 1000000).toFixed(0)}Jt`;
        else budgetStr = `${budget.toLocaleString()}`;
      } else if (isReach) {
        if (actual >= 1000000) actualStr = `${(actual / 1000000).toFixed(1)}M`;
        else if (actual >= 1000) actualStr = `${(actual / 1000).toFixed(1)}K`;
        else actualStr = `${actual.toLocaleString()}`;

        if (budget >= 1000000) budgetStr = `${(budget / 1000000).toFixed(1)}M`;
        else if (budget >= 1000) budgetStr = `${(budget / 1000).toFixed(1)}K`;
        else budgetStr = `${budget.toLocaleString()}`;
      } else {
        actualStr = `${actual}`;
        budgetStr = `${budget}`;
      }

      const remaining = Math.max(0, budget - actual);
      const chartData = [
        { name: "Actual", value: actual, fill: name === "TOTAL" ? "#f59e0b" : "#06b6d4" },
        { name: "Remaining", value: remaining, fill: "#f1f5f9" }
      ];

      if (actual > budget) {
        chartData[0].value = actual;
        chartData[1].value = 0;
      }

      return {
        name,
        fullName: name === "FFD" ? "Farmer Field Day" :
                  name === "FM" ? "Farmer Meeting" :
                  name === "ODP" ? "One Day Promo" :
                  name === "SFT" ? "Special Field Trip" :
                  name === "BFFD" ? "Big Farmer Field Day" :
                  name === "BFM" ? "Big Farmer Meeting" :
                  name === "BFT" ? "Big Field Trip" :
                  name === "BC" ? "Branding Crop" :
                  name === "CRV" ? "Caravan" :
                  name === "EXP" ? "Expo" :
                  name === "PT" ? "Pasar Tani" : "Total Activity",
        budget,
        actual,
        actualStr,
        budgetStr,
        percentage: pct,
        chartData
      };
    });
  }, [pogDataOverviewFiltered, overviewMetricFilter, filterBelowType]);

  const overviewStats = useMemo(() => {
    let activeKiosks = kiosks || [];
    if (filterBelowChannel && filterBelowChannel !== "All") {
      activeKiosks = activeKiosks.filter(
        (k) => cleanForMatch(k.category) === cleanForMatch(filterBelowChannel),
      );
    }
    if (filterBelowTeam && filterBelowTeam !== "All") {
      activeKiosks = activeKiosks.filter((k) => {
        const resolvedPic = getDdaOfUser(
          k.pic || "",
          userData?.name,
          computedTeamProfiles,
        );
        return cleanForMatch(resolvedPic) === cleanForMatch(filterBelowTeam);
      });
    }
    if (filterBelowArea && filterBelowArea !== "All") {
      activeKiosks = activeKiosks.filter((k) => {
        const resolvedPic = getDdaOfUser(
          k.pic || "",
          userData?.name,
          computedTeamProfiles,
        );
        const matchedMember = teamMembers.find((m) =>
          matchNames(m, resolvedPic),
        );
        const area =
          getFromRecord<string>(teamAreas, matchedMember || resolvedPic) ||
          k.area ||
          "-";
        return cleanForMatch(area) === cleanForMatch(filterBelowArea);
      });
    }
    if (filterBelowCrop && filterBelowCrop !== "All") {
      const activeKioskNames = new Set(
        pogDataOverviewFiltered.map((item) => cleanForMatch(item.kiosk)),
      );
      activeKiosks = activeKiosks.filter((k) =>
        activeKioskNames.has(cleanForMatch(k.name)),
      );
    }

    let totalKiosks = activeKiosks.length;
    let totalOpeningStock = 0;
    let totalSellIn = 0;
    let totalSellOut = 0;
    let totalCurrentStock = 0;
    let totalIdleStock = 0;

    pogDataOverviewFiltered.forEach((item) => {
      totalOpeningStock += Number(item.lastQty || 0);
      totalSellIn += Number(item.sellIn || 0);
      totalSellOut += Number(item.sellOut || 0);
      totalCurrentStock += Number(item.currentQty || 0);
      totalIdleStock += Number(item.idleStock || 0);
    });

    const AREAS_LIST = ["T1", "T2", "T3", "T4", "T5", "T6", "T7"];
    const PROVINCES_LIST = [
      "North Sumatra", "West Sumatra", "Lampung",
      "Central Java", "East Java", "Gorontalo",
      "North Sulawesi", "South Sulawesi", "NTB"
    ];
    const SALES_AGRONOMISTS_LIST = [
      "Rian Hidayat", "Agus Hermawan", "Dedi Setiawan", "Budi Santoso",
      "Hendra Kurniawan", "Eko Prasetyo", "Andi Wijaya", "Slamet Riyadi",
      "Yanto Subagyo", "Feri Nugroho", "Joko Susilo", "Rudi Hartono",
      "Mulyono", "Setiawan", "Herianto", "Bambang", "Edi Purwanto", "Wawan"
    ];
    const HYBRIDS_LIST = ["JAGO", "RUBY", "JALU", "GANESH"];
    const ACTIVITIES_LIST = ["FFD", "FM", "ODP", "SFT", "BFFD", "BFM", "BFT", "BC", "CRV", "EXP", "PT"];

    const getDeterministicDummy = (name: string, index: number) => {
      // Create high-quality, realistic dummy metrics based on the name hash
      const charSum = name.split("").reduce((sum, ch) => sum + ch.charCodeAt(0), 0);
      const seedVal = (charSum + index * 47) % 100;
      const multiplier = 0.7 + (seedVal / 100) * 0.8; // between 0.7 and 1.5

      const opening = Math.floor((1500 + (seedVal * 17) % 1000) * multiplier);
      const sellIn = Math.floor((1200 + (seedVal * 31) % 800) * multiplier);
      const pog = Math.floor((1100 + (seedVal * 23) % 900) * multiplier);
      const sellOut = Math.floor(pog * 1.08);
      const stock = Math.max(150, opening + sellIn - sellOut);
      const idle = Math.floor(stock * 0.14);

      // Activity metrics: budget vs actual (e.g. 10 to 80 events)
      const budgetActivity = Math.floor((15 + (seedVal * 7) % 65) * multiplier);
      const actualActivity = Math.floor(budgetActivity * (0.8 + (seedVal % 20) / 100));

      // Nominal metrics: budget vs actual (e.g. Rp 50.000.000 to Rp 350.000.000)
      const budgetNominal = Math.floor((50000000 + (seedVal * 3500000) % 300000000) * multiplier);
      const actualNominal = Math.floor(budgetNominal * (0.82 + (seedVal % 18) / 100));

      // Reach metrics: number of attending farmers (e.g. 300 to 2400 farmers)
      const budgetReach = Math.floor(budgetActivity * (20 + (seedVal % 15)));
      const actualReach = Math.floor(actualActivity * (18 + ((seedVal + 3) % 15)));

      return {
        name,
        pog,
        stock,
        sellIn,
        sellOut,
        opening,
        idle,
        budgetActivity,
        actualActivity,
        budgetNominal,
        actualNominal,
        budgetReach,
        actualReach,
      };
    };

    // Generic helper to build grouped chart data by dimension
    const buildChartData = (dimension: string, filterByMainKey?: string | null) => {
      const gMap: Record<
        string,
        {
          name: string;
          pog: number;
          stock: number;
          sellIn: number;
          sellOut: number;
          opening: number;
          idle: number;
          budgetActivity: number;
          actualActivity: number;
          budgetNominal: number;
          actualNominal: number;
          budgetReach: number;
          actualReach: number;
        }
      > = {};

      const activityScale = activeActivityFilter ? 0.25 : 1.0;
      const scaleFactor = (filterByMainKey ? 0.15 : 1.0) * activityScale;

      if (dimension === "area") {
        AREAS_LIST.forEach((area, idx) => {
          const dummy = getDeterministicDummy(area, idx);
          if (filterByMainKey || activeActivityFilter) {
            dummy.pog = Math.round(dummy.pog * scaleFactor);
            dummy.stock = Math.round(dummy.stock * scaleFactor);
            dummy.sellIn = Math.round(dummy.sellIn * scaleFactor);
            dummy.sellOut = Math.round(dummy.sellOut * scaleFactor);
            dummy.opening = Math.round(dummy.opening * scaleFactor);
            dummy.idle = Math.round(dummy.idle * scaleFactor);
            dummy.budgetActivity = Math.round(dummy.budgetActivity * scaleFactor);
            dummy.actualActivity = Math.round(dummy.actualActivity * scaleFactor);
            dummy.budgetNominal = Math.round(dummy.budgetNominal * scaleFactor);
            dummy.actualNominal = Math.round(dummy.actualNominal * scaleFactor);
            dummy.budgetReach = Math.round(dummy.budgetReach * scaleFactor);
            dummy.actualReach = Math.round(dummy.actualReach * scaleFactor);
          }
          gMap[area] = dummy;
        });
      } else if (dimension === "province") {
        PROVINCES_LIST.forEach((prov, idx) => {
          const dummy = getDeterministicDummy(prov, idx);
          if (filterByMainKey || activeActivityFilter) {
            dummy.pog = Math.round(dummy.pog * scaleFactor);
            dummy.stock = Math.round(dummy.stock * scaleFactor);
            dummy.sellIn = Math.round(dummy.sellIn * scaleFactor);
            dummy.sellOut = Math.round(dummy.sellOut * scaleFactor);
            dummy.opening = Math.round(dummy.opening * scaleFactor);
            dummy.idle = Math.round(dummy.idle * scaleFactor);
            dummy.budgetActivity = Math.round(dummy.budgetActivity * scaleFactor);
            dummy.actualActivity = Math.round(dummy.actualActivity * scaleFactor);
            dummy.budgetNominal = Math.round(dummy.budgetNominal * scaleFactor);
            dummy.actualNominal = Math.round(dummy.actualNominal * scaleFactor);
            dummy.budgetReach = Math.round(dummy.budgetReach * scaleFactor);
            dummy.actualReach = Math.round(dummy.actualReach * scaleFactor);
          }
          gMap[prov] = dummy;
        });
      } else if (dimension === "sales_agronomist") {
        SALES_AGRONOMISTS_LIST.forEach((sa, idx) => {
          const dummy = getDeterministicDummy(sa, idx);
          if (filterByMainKey || activeActivityFilter) {
            dummy.pog = Math.round(dummy.pog * scaleFactor);
            dummy.stock = Math.round(dummy.stock * scaleFactor);
            dummy.sellIn = Math.round(dummy.sellIn * scaleFactor);
            dummy.sellOut = Math.round(dummy.sellOut * scaleFactor);
            dummy.opening = Math.round(dummy.opening * scaleFactor);
            dummy.idle = Math.round(dummy.idle * scaleFactor);
            dummy.budgetActivity = Math.round(dummy.budgetActivity * scaleFactor);
            dummy.actualActivity = Math.round(dummy.actualActivity * scaleFactor);
            dummy.budgetNominal = Math.round(dummy.budgetNominal * scaleFactor);
            dummy.actualNominal = Math.round(dummy.actualNominal * scaleFactor);
            dummy.budgetReach = Math.round(dummy.budgetReach * scaleFactor);
            dummy.actualReach = Math.round(dummy.actualReach * scaleFactor);
          }
          gMap[sa] = dummy;
        });
      } else if (dimension === "hybrid" || dimension === "material") {
        HYBRIDS_LIST.forEach((hyb, idx) => {
          const dummy = getDeterministicDummy(hyb, idx);
          if (filterByMainKey || activeActivityFilter) {
            dummy.pog = Math.round(dummy.pog * scaleFactor);
            dummy.stock = Math.round(dummy.stock * scaleFactor);
            dummy.sellIn = Math.round(dummy.sellIn * scaleFactor);
            dummy.sellOut = Math.round(dummy.sellOut * scaleFactor);
            dummy.opening = Math.round(dummy.opening * scaleFactor);
            dummy.idle = Math.round(dummy.idle * scaleFactor);
            dummy.budgetActivity = Math.round(dummy.budgetActivity * scaleFactor);
            dummy.actualActivity = Math.round(dummy.actualActivity * scaleFactor);
            dummy.budgetNominal = Math.round(dummy.budgetNominal * scaleFactor);
            dummy.actualNominal = Math.round(dummy.actualNominal * scaleFactor);
            dummy.budgetReach = Math.round(dummy.budgetReach * scaleFactor);
            dummy.actualReach = Math.round(dummy.actualReach * scaleFactor);
          }
          gMap[hyb] = dummy;
        });
      } else if (dimension === "activity") {
        ACTIVITIES_LIST.forEach((act, idx) => {
          const dummy = getDeterministicDummy(act, idx);
          if (filterByMainKey || activeActivityFilter) {
            dummy.pog = Math.round(dummy.pog * scaleFactor);
            dummy.stock = Math.round(dummy.stock * scaleFactor);
            dummy.sellIn = Math.round(dummy.sellIn * scaleFactor);
            dummy.sellOut = Math.round(dummy.sellOut * scaleFactor);
            dummy.opening = Math.round(dummy.opening * scaleFactor);
            dummy.idle = Math.round(dummy.idle * scaleFactor);
            dummy.budgetActivity = Math.round(dummy.budgetActivity * scaleFactor);
            dummy.actualActivity = Math.round(dummy.actualActivity * scaleFactor);
            dummy.budgetNominal = Math.round(dummy.budgetNominal * scaleFactor);
            dummy.actualNominal = Math.round(dummy.actualNominal * scaleFactor);
            dummy.budgetReach = Math.round(dummy.budgetReach * scaleFactor);
            dummy.actualReach = Math.round(dummy.actualReach * scaleFactor);
          }
          gMap[act] = dummy;
        });
      }

      pogDataOverviewFiltered.forEach((item) => {
        // If activeActivityFilter is active, filter items by it!
        const kioskCharSumForAct = (item.kiosk || "").split("").reduce((sum, ch) => sum + ch.charCodeAt(0), 0);
        const actIdxForAct = kioskCharSumForAct % ACTIVITIES_LIST.length;
        const itemAct = ACTIVITIES_LIST[actIdxForAct];

        if (activeActivityFilter && itemAct !== activeActivityFilter) {
          return;
        }

        // If filterByMainKey is active, verify if item belongs to that main key
        if (filterByMainKey) {
          const itemPic = item.pic;
          const rawPos = getFromRecord<string>(teamPositions, itemPic) || "";
          const pos = normalizePosition(rawPos);
          let belongs = false;

          if (overviewGroupDimension === "area") {
            const itemArea = item.area || "";
            belongs = cleanForMatch(itemArea) === cleanForMatch(filterByMainKey);
          } else if (overviewGroupDimension === "province") {
            const itemProv = getFromRecord<string>(teamProvinces, itemPic) || item.province || "";
            belongs = cleanForMatch(itemProv) === cleanForMatch(filterByMainKey);
          } else if (overviewGroupDimension === "sales_agronomist") {
            if (pos === "Sales Agronomist" && itemPic) {
              belongs = cleanForMatch(itemPic) === cleanForMatch(filterByMainKey);
            }
          } else if (overviewGroupDimension === "hybrid" || overviewGroupDimension === "material") {
            const itemHyb = item.hybrid || "";
            belongs = cleanForMatch(itemHyb) === cleanForMatch(filterByMainKey);
          } else if (overviewGroupDimension === "activity") {
            const kioskCharSum = (item.kiosk || "").split("").reduce((sum, ch) => sum + ch.charCodeAt(0), 0);
            const actIdx = kioskCharSum % ACTIVITIES_LIST.length;
            const act = ACTIVITIES_LIST[actIdx];
            belongs = cleanForMatch(act) === cleanForMatch(filterByMainKey);
          }

          if (!belongs) return;
        }

        let gVal = "";
        const itemPic = item.pic;
        const rawPos = getFromRecord<string>(teamPositions, itemPic) || "";
        const pos = normalizePosition(rawPos);

        if (dimension === "area") {
          const itemArea = item.area || "";
          const matched = AREAS_LIST.find(a => cleanForMatch(a) === cleanForMatch(itemArea));
          if (matched) gVal = matched;
        } else if (dimension === "province") {
          const itemProv = getFromRecord<string>(teamProvinces, itemPic) || item.province || "";
          const matched = PROVINCES_LIST.find(p => cleanForMatch(p) === cleanForMatch(itemProv));
          if (matched) gVal = matched;
        } else if (dimension === "sales_agronomist") {
          if (pos === "Sales Agronomist" && itemPic) {
            const matched = SALES_AGRONOMISTS_LIST.find(sa => cleanForMatch(sa) === cleanForMatch(itemPic));
            if (matched) gVal = matched;
          }
        } else if (dimension === "hybrid" || dimension === "material") {
          const itemHyb = item.hybrid || "";
          const matched = HYBRIDS_LIST.find(h => cleanForMatch(h) === cleanForMatch(itemHyb));
          if (matched) gVal = matched;
        } else if (dimension === "activity") {
          const kioskCharSum = (item.kiosk || "").split("").reduce((sum, ch) => sum + ch.charCodeAt(0), 0);
          const actIdx = kioskCharSum % ACTIVITIES_LIST.length;
          gVal = ACTIVITIES_LIST[actIdx];
        }

        if (gVal) {
          if (!gMap[gVal]) {
            gMap[gVal] = {
              name: gVal,
              pog: 0,
              stock: 0,
              sellIn: 0,
              sellOut: 0,
              opening: 0,
              idle: 0,
              budgetActivity: 0,
              actualActivity: 0,
              budgetNominal: 0,
              actualNominal: 0,
              budgetReach: 0,
              actualReach: 0,
            };
          }
          gMap[gVal].pog += Number(item.pog || 0);
          gMap[gVal].stock += Number(item.currentQty || 0);
          gMap[gVal].sellIn += Number(item.sellIn || 0);
          gMap[gVal].sellOut += Number(item.sellOut || 0);
          gMap[gVal].opening += Number(item.lastQty || 0);
          gMap[gVal].idle += Number(item.idleStock || 0);
        }
      });

      return Object.values(gMap).sort((a, b) => {
        if (overviewMetricFilter === "activity") {
          return b.budgetActivity - a.budgetActivity;
        } else if (overviewMetricFilter === "nominal") {
          return b.budgetNominal - a.budgetNominal;
        } else if (overviewMetricFilter === "reach") {
          return b.budgetReach - a.budgetReach;
        } else if (overviewMetricFilter === "movement") {
          return b.sellIn + b.pog - (a.sellIn + a.pog);
        } else if (overviewMetricFilter === "idle") {
          return b.idle - a.idle;
        } else if (overviewMetricFilter === "total_stock") {
          return b.stock - a.stock;
        } else if (overviewMetricFilter === "Opening") {
          return b.opening - a.opening;
        } else {
          return b.pog - a.pog;
        }
      });
    };

    const areaChartData = buildChartData(overviewGroupDimension);
    const subChartData = buildChartData(subGroupDimension, activeMainBarKey);

    // Group by Category of Kiosk (Channel)
    const catMap: Record<string, { name: string; value: number }> = {};
    activeKiosks.forEach((k) => {
      const cat = k.category || "Uncategorized";
      if (!catMap[cat]) {
        catMap[cat] = { name: cat, value: 0 };
      }
      catMap[cat].value += 1;
    });
    const categoryChartData = Object.values(catMap);

    // Group by selected partnerSegmentDimension
    const partnerMap: Record<string, { name: string; actual: number; budget: number }> = {};
    const SEG_AREAS_LIST = ["T1", "T2", "T3", "T4", "T5", "T6", "T7"];
    const SEG_PROVINCES_LIST = [
      "North Sumatra", "West Sumatra", "Lampung",
      "Central Java", "East Java", "Gorontalo",
      "North Sulawesi", "South Sulawesi", "NTB"
    ];
    const SEG_SALES_AGRONOMISTS_LIST = [
      "Rian Hidayat", "Agus Hermawan", "Dedi Setiawan", "Budi Santoso",
      "Hendra Kurniawan", "Eko Prasetyo", "Andi Wijaya", "Slamet Riyadi",
      "Yanto Subagyo", "Feri Nugroho", "Joko Susilo", "Rudi Hartono",
      "Mulyono", "Setiawan", "Herianto", "Bambang", "Edi Purwanto", "Wawan"
    ];
    const SEG_HYBRIDS_LIST = ["JAGO", "RUBY", "JALU", "GANESH"];
    const SEG_ACTIVITIES_LIST = ["FFD", "FM", "ODP", "SFT", "BFFD", "BFM", "BFT", "BC", "CRV", "EXP", "PT"];

    activeKiosks.forEach((k) => {
      let groupName = "Uncategorized";
      const charSum = k.name.split("").reduce((sum, ch) => sum + ch.charCodeAt(0), 0);

      if (partnerSegmentDimension === "area") {
        groupName = k.area || "Lainnya";
        const matched = SEG_AREAS_LIST.find(a => cleanForMatch(a) === cleanForMatch(groupName));
        if (matched) groupName = matched;
      } else if (partnerSegmentDimension === "province") {
        groupName = k.province || "Lainnya";
        const matched = SEG_PROVINCES_LIST.find(p => cleanForMatch(p) === cleanForMatch(groupName));
        if (matched) groupName = matched;
      } else if (partnerSegmentDimension === "sales_agronomist") {
        groupName = k.pic || "Lainnya";
        const matched = SEG_SALES_AGRONOMISTS_LIST.find(sa => cleanForMatch(sa) === cleanForMatch(groupName));
        if (matched) groupName = matched;
      } else if (partnerSegmentDimension === "hybrid") {
        const hybIdx = charSum % SEG_HYBRIDS_LIST.length;
        groupName = SEG_HYBRIDS_LIST[hybIdx];
      } else if (partnerSegmentDimension === "activity") {
        const actIdx = charSum % SEG_ACTIVITIES_LIST.length;
        groupName = SEG_ACTIVITIES_LIST[actIdx];
      }

      if (!partnerMap[groupName]) {
        partnerMap[groupName] = { name: groupName, actual: 0, budget: 0 };
      }
      partnerMap[groupName].actual += 1;
    });

    // Populate budget deterministically based on actual
    Object.keys(partnerMap).forEach((key) => {
      const act = partnerMap[key].actual;
      partnerMap[key].budget = Math.round(act * 1.2) + 2;
    });

    let partnerChartData: Array<{ name: string; actual: number; budget: number }> = Object.values(partnerMap);
    if (partnerChartData.length === 0) {
      if (partnerSegmentDimension === "area") {
        partnerChartData = [
          { name: "T1", actual: 48, budget: 55 },
          { name: "T2", actual: 39, budget: 45 },
          { name: "T3", actual: 30, budget: 35 },
          { name: "T4", actual: 22, budget: 25 },
          { name: "T5", actual: 18, budget: 22 },
          { name: "T6", actual: 15, budget: 18 },
          { name: "T7", actual: 12, budget: 15 }
        ];
      } else if (partnerSegmentDimension === "province") {
        partnerChartData = [
          { name: "North Sumatra", actual: 42, budget: 48 },
          { name: "Central Java", actual: 31, budget: 35 },
          { name: "East Java", actual: 25, budget: 30 },
          { name: "South Sulawesi", actual: 19, budget: 25 },
          { name: "NTB", actual: 14, budget: 18 },
          { name: "Lampung", actual: 8, budget: 10 }
        ];
      } else if (partnerSegmentDimension === "sales_agronomist") {
        partnerChartData = [
          { name: "Rian Hidayat", actual: 35, budget: 40 },
          { name: "Agus Hermawan", actual: 28, budget: 32 },
          { name: "Dedi Setiawan", actual: 24, budget: 28 },
          { name: "Budi Santoso", actual: 20, budget: 24 },
          { name: "Hendra Kurniawan", actual: 15, budget: 18 }
        ];
      } else if (partnerSegmentDimension === "hybrid") {
        partnerChartData = [
          { name: "JAGO", actual: 55, budget: 60 },
          { name: "RUBY", actual: 42, budget: 48 },
          { name: "JALU", actual: 38, budget: 42 },
          { name: "GANESH", actual: 29, budget: 35 }
        ];
      } else if (partnerSegmentDimension === "activity") {
        partnerChartData = [
          { name: "FFD", actual: 64, budget: 70 },
          { name: "FM", actual: 48, budget: 55 },
          { name: "ODP", actual: 36, budget: 40 },
          { name: "SFT", actual: 22, budget: 25 },
          { name: "BFFD", actual: 18, budget: 20 },
          { name: "BFM", actual: 32, budget: 35 },
          { name: "BFT", actual: 12, budget: 15 },
          { name: "BC", actual: 50, budget: 60 },
          { name: "CRV", actual: 28, budget: 30 },
          { name: "EXP", actual: 10, budget: 12 },
          { name: "PT", actual: 45, budget: 50 }
        ];
      }
    }

    // Sort descending by actual
    partnerChartData.sort((a, b) => b.actual - a.actual);

    // Group by Crops
    const cropMap: Record<
      string,
      { name: string; stock: number; pog: number }
    > = {};
    pogDataOverviewFiltered.forEach((item) => {
      const crop = item.crops || "Lainnya";
      if (!cropMap[crop]) {
        cropMap[crop] = { name: crop, stock: 0, pog: 0 };
      }
      cropMap[crop].stock += Number(item.currentQty || 0);
      cropMap[crop].pog += Number(item.pog || 0);
    });
    const cropsChartData = Object.values(cropMap).sort(
      (a, b) => b.stock - a.stock,
    );

    return {
      totalKiosks,
      totalOpeningStock,
      totalSellIn,
      totalSellOut,
      totalCurrentStock,
      totalIdleStock,
      areaChartData,
      subChartData,
      categoryChartData,
      cropsChartData,
      partnerChartData,
    };
  }, [
    kiosks,
    pogDataOverviewFiltered,
    overviewMetricFilter,
    overviewGroupDimension,
    subGroupDimension,
    partnerSegmentDimension,
    teamPositions,
    teamProvinces,
    filterBelowChannel,
    filterBelowTeam,
    filterBelowArea,
    filterBelowCrop,
    userData,
    teamMembers,
    teamAreas,
    activeMainBarKey,
    activeActivityFilter,
  ]);

  const topKiosksData = useMemo(() => {
    const kioskSales: Record<
      string,
      {
        kiosk: string;
        category: string;
        pic: string;
        area: string;
        pog: number;
        currentStock: number;
      }
    > = {};

    pogDataOverviewFiltered.forEach((item) => {
      const kName = item.kiosk;
      if (!kioskSales[kName]) {
        kioskSales[kName] = {
          kiosk: kName,
          category: item.category || "Uncategorized",
          pic: item.pic || "Unknown",
          area: item.area || "-",
          pog: 0,
          currentStock: 0,
        };
      }
      kioskSales[kName].pog += Number(item.pog || 0);
      kioskSales[kName].currentStock += Number(item.currentQty || 0);
    });

    return Object.values(kioskSales)
      .sort((a, b) => b.pog - a.pog)
      .slice(0, 5);
  }, [pogDataOverviewFiltered]);

  const employeePerformanceData = useMemo(() => {
    const empSales: Record<
      string,
      { employee: string; area: string; pog: number; currentStock: number }
    > = {};

    teamMembers.forEach((member) => {
      const mClean = cleanForMatch(member);

      // Get area for this member
      let area = "-";
      const foundArea = getFromRecord<string>(teamAreas, member);
      if (foundArea) {
        area = foundArea;
      } else if (mClean === cleanForMatch(userData?.name)) {
        area = userData?.area || "-";
      }

      // Filter by Area if chosen
      if (
        filterBelowArea &&
        filterBelowArea !== "All" &&
        cleanForMatch(area) !== cleanForMatch(filterBelowArea)
      ) {
        return;
      }

      // Filter by PIC if chosen
      if (
        filterBelowTeam &&
        filterBelowTeam !== "All" &&
        mClean !== cleanForMatch(filterBelowTeam)
      ) {
        return;
      }

      empSales[mClean] = {
        employee: member,
        area: area,
        pog: 0,
        currentStock: 0,
      };
    });

    // Accumulate sales and stock from pogDataOverviewFiltered
    pogDataOverviewFiltered.forEach((item) => {
      const picClean = cleanForMatch(item.pic);
      // Try to find matching employee
      let matchedEmpClean = picClean;
      if (!empSales[picClean]) {
        const found = teamMembers.find(
          (m) => cleanForMatch(m) === picClean || matchNames(m, item.pic),
        );
        if (found) {
          matchedEmpClean = cleanForMatch(found);
        }
      }

      if (empSales[matchedEmpClean]) {
        empSales[matchedEmpClean].pog += Number(item.pog || 0);
        empSales[matchedEmpClean].currentStock += Number(item.currentQty || 0);
      }
    });

    const empSalesList = Object.values(empSales);

    // Sort for Highest (Top 5)
    const highest = [...empSalesList].sort((a, b) => b.pog - a.pog).slice(0, 5);

    // Sort for Lowest (Bottom 5)
    const lowest = [...empSalesList].sort((a, b) => a.pog - b.pog).slice(0, 5);

    return { highest, lowest };
  }, [
    teamMembers,
    pogDataOverviewFiltered,
    teamAreas,
    userData,
    filterBelowArea,
    filterBelowTeam,
  ]);

  // Filter active POG team members based on selected crop filter
  const activePogMembers = useMemo(() => {
    if (!filterBelowCrop || filterBelowCrop === "All") return teamMembers;

    const kiosksMapByCleanName: Record<string, any> = {};
    kiosks.forEach((k) => {
      kiosksMapByCleanName[cleanForMatch(k.name)] = k;
    });
    const teamMembersMapByCleanName: Record<string, string> = {};
    teamMembers.forEach((m) => {
      teamMembersMapByCleanName[cleanForMatch(m)] = m;
    });
    const getTeamMemberMatch = (name: string): string | undefined => {
      const clean = cleanForMatch(name);
      if (teamMembersMapByCleanName[clean])
        return teamMembersMapByCleanName[clean];
      return teamMembers.find((m) => matchNames(m, name));
    };

    const matchedMembersWithCrop = new Set<string>();
    pogDataProcessed.forEach((item) => {
      const itemCrop = String(item.crops || "")
        .trim()
        .toLowerCase();
      if (checkCropMatch(itemCrop, filterBelowCrop)) {
        const matchedMember = getTeamMemberMatch(item.pic);
        if (matchedMember) {
          matchedMembersWithCrop.add(cleanForMatch(matchedMember));
        }
      }
    });

    const activeSet = new Set<string>();
    activeSet.add(cleanForMatch(userData?.name));

    teamMembers.forEach((m) => {
      const mClean = cleanForMatch(m);
      if (matchedMembersWithCrop.has(mClean)) {
        activeSet.add(mClean);
        let current = m;
        for (let i = 0; i < 15; i++) {
          const upRaw = getFromRecord<string>(teamUpLines, current);
          if (!upRaw) break;

          const up = getTeamMemberMatch(upRaw as string) || (upRaw as string);

          const upClean = cleanForMatch(up);
          if (activeSet.has(upClean)) {
            break;
          }
          activeSet.add(upClean);
          current = up;
        }
      }
    });

    return teamMembers.filter((m) => activeSet.has(cleanForMatch(m)));
  }, [
    teamMembers,
    filterBelowCrop,
    pogDataProcessed,
    kiosks,
    userData,
    teamUpLines,
  ]);

  const aggregatedPogData = useMemo(() => {
    const teamMembersCleanSet = new Set(
      activePogMembers.map((m) => cleanForMatch(m)).filter(Boolean),
    );
    

    const filteredData = pogDataProcessed.filter((item) => {
      const picClean = cleanForMatch(item.pic);
      const isTeamMember =
        teamMembersCleanSet.has(picClean);
      const isCropMatch = checkCropMatch(item.crops, filterBelowCrop);
      return isTeamMember && isCropMatch;
    });

    // Filter active POG team members based on selected crop filter
    const hasGroupInfo = employees.some((e) => {
      const g = String(
        e.group || e.Group || e["group"] || e["Group"] || "",
      ).trim();
      return g.length > 0;
    });

    if (pogGroupBy === "subordinate") {
      // Pre-compute uplines for activePogMembers to avoid O(M) recursive scans inside buildNode loops
      const pMembersUplines: Record<string, string | null> = {};
      const pDirectSubsMap: Record<string, string[]> = {};
      activePogMembers.forEach((m) => {
        pMembersUplines[m] = getUplineInTeam(m, activePogMembers, teamUpLines);
      });

      const rootMembers = activePogMembers.filter((m) => {
        return pMembersUplines[m] === null;
      });
      rootMembers.sort((a, b) =>
        compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
      );

      activePogMembers.forEach((m) => {
        const upl = pMembersUplines[m];
        if (upl !== null) {
          const uplClean = cleanForMatch(upl);
          if (!pDirectSubsMap[uplClean]) {
            pDirectSubsMap[uplClean] = [];
          }
          pDirectSubsMap[uplClean].push(m);
        }
      });

      Object.keys(pDirectSubsMap).forEach((key) => {
        pDirectSubsMap[key].sort((a, b) =>
          compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
        );
      });

      const buildNode = (name: string): any => {
        const nameClean = cleanForMatch(name);
        const directSubs = pDirectSubsMap[nameClean] || [];
        const myItems = filteredData.filter(
          (item) => cleanForMatch(item.pic) === nameClean,
        );
        const childrenNodes = directSubs.map((sub) => buildNode(sub));

        const ownPog = {
          lastQty: 0,
          currentQty: 0,
          pog: 0,
          sellIn: 0,
          sellOut: 0,
          totalInv: 0,
          idleStock: 0,
        };
        myItems.forEach((item) => {
          ownPog.lastQty += Number(item.lastQty) || 0;
          ownPog.currentQty += Number(item.currentQty) || 0;
          ownPog.sellIn += Number(item.sellIn) || 0;
          ownPog.sellOut += Number(item.sellOut) || 0;
          ownPog.totalInv += Number(item.totalInv) || 0;
          ownPog.pog += Number(item.pog) || 0;
          ownPog.idleStock += Number(item.idleStock) || 0;
        });

        const transPog = { ...ownPog };
        childrenNodes.forEach((child) => {
          transPog.lastQty += child.lastQty || 0;
          transPog.currentQty += child.currentQty || 0;
          transPog.sellIn += child.sellIn || 0;
          transPog.sellOut += child.sellOut || 0;
          transPog.totalInv += child.totalInv || 0;
          transPog.pog += child.pog || 0;
          transPog.idleStock += child.idleStock || 0;
        });

        let finalChildren = [];

        if (pogSubGroupBy === "subordinate") {
          finalChildren = [...childrenNodes];
        } else {
          const leafGroups: Record<string, any> = {};

          // Accumulate own items
          myItems.forEach((item) => {
            let leafKey = "Unknown";
            if (pogSubGroupBy === "channel" || pogSubGroupBy === "kiosk")
              leafKey = item.kiosk || "Unknown Channel";
            else if (pogSubGroupBy === "hybrid")
              leafKey = item.hybrid || "Unknown";
            else if (pogSubGroupBy === "area") leafKey = item.area || "Unknown";
            else if (pogSubGroupBy === "category")
              leafKey = item.category || "Unknown";
            else if (pogSubGroupBy === "crops")
              leafKey = item.crops || "Unknown";

            if (!leafGroups[leafKey]) {
              leafGroups[leafKey] = {
                lastQty: 0,
                currentQty: 0,
                pog: 0,
                sellIn: 0,
                sellOut: 0,
                totalInv: 0,
                idleStock: 0,
                category: item.category,
              };
            }
            leafGroups[leafKey].lastQty += Number(item.lastQty) || 0;
            leafGroups[leafKey].currentQty += Number(item.currentQty) || 0;
            leafGroups[leafKey].sellIn += Number(item.sellIn) || 0;
            leafGroups[leafKey].sellOut += Number(item.sellOut) || 0;
            leafGroups[leafKey].totalInv += Number(item.totalInv) || 0;
            leafGroups[leafKey].pog += Number(item.pog) || 0;
            leafGroups[leafKey].idleStock += Number(item.idleStock) || 0;
          });

          // Roll up children's leaves
          childrenNodes.forEach((child) => {
            (child.children || []).forEach((cNode: any) => {
              const leafKey = cNode.name;
              if (!leafGroups[leafKey]) {
                leafGroups[leafKey] = {
                  lastQty: 0,
                  currentQty: 0,
                  pog: 0,
                  sellIn: 0,
                  sellOut: 0,
                  totalInv: 0,
                  idleStock: 0,
                  category: cNode.category,
                };
              }
              leafGroups[leafKey].lastQty += Number(cNode.lastQty) || 0;
              leafGroups[leafKey].currentQty += Number(cNode.currentQty) || 0;
              leafGroups[leafKey].sellIn += Number(cNode.sellIn) || 0;
              leafGroups[leafKey].sellOut += Number(cNode.sellOut) || 0;
              leafGroups[leafKey].totalInv += Number(cNode.totalInv) || 0;
              leafGroups[leafKey].pog += Number(cNode.pog) || 0;
              leafGroups[leafKey].idleStock += Number(cNode.idleStock) || 0;
            });
          });

          let leaves = Object.entries(leafGroups).map(([leafName, val]) => ({
            name: leafName,
            isLeaf: true,
            ...(val as any),
          }));

          if (pogSubGroupBy === "channel" || pogSubGroupBy === "kiosk") {
            const catOrder: Record<string, number> = {
              Distributor: 1,
              R1: 2,
              R2: 3,
            };
            leaves.sort((a, b) => {
              const wA = catOrder[a.category] || 99;
              const wB = catOrder[b.category] || 99;
              if (wA !== wB) return wA - wB;
              return b.pog - a.pog;
            });
          } else {
            leaves.sort((a, b) => b.pog - a.pog);
          }

          finalChildren = leaves;
        }

        return {
          name,
          level: getMemberLevel(name, teamLevels, teamPositions, userData),
          children: finalChildren,
          isExpandable: finalChildren.length > 0,
          teamChildren: childrenNodes,
          ...transPog,
        };
      };

      const roots = rootMembers.map((root) => buildNode(root));
      roots.sort((a, b) => b.pog - a.pog);
      return roots.flatMap((node) => {
        if (cleanForMatch(node.name) === cleanForMatch(userData?.name)) {
          if (userLevel <= 3) {
            const nameClean = cleanForMatch(node.name);
            const myItems = filteredData.filter(
              (item) => cleanForMatch(item.pic) === nameClean,
            );

            const ownPog = {
              lastQty: 0,
              currentQty: 0,
              pog: 0,
              sellIn: 0,
              sellOut: 0,
              totalInv: 0,
              idleStock: 0,
            };
            myItems.forEach((item) => {
              ownPog.lastQty += Number(item.lastQty) || 0;
              ownPog.currentQty += Number(item.currentQty) || 0;
              ownPog.sellIn += Number(item.sellIn) || 0;
              ownPog.sellOut += Number(item.sellOut) || 0;
              ownPog.totalInv += Number(item.totalInv) || 0;
              ownPog.pog += Number(item.pog) || 0;
              ownPog.idleStock += Number(item.idleStock) || 0;
            });

            let finalChildrenOfSelf = [];
            if (pogSubGroupBy !== "subordinate") {
              const leafGroups: Record<string, any> = {};
              myItems.forEach((item) => {
                let leafKey = "Unknown";
                if (pogSubGroupBy === "channel" || pogSubGroupBy === "kiosk")
                  leafKey = item.kiosk || "Unknown Channel";
                else if (pogSubGroupBy === "hybrid")
                  leafKey = item.hybrid || "Unknown";
                else if (pogSubGroupBy === "area")
                  leafKey = item.area || "Unknown";
                else if (pogSubGroupBy === "category")
                  leafKey = item.category || "Unknown";
                else if (pogSubGroupBy === "crops")
                  leafKey = item.crops || "Unknown";

                if (!leafGroups[leafKey]) {
                  leafGroups[leafKey] = {
                    lastQty: 0,
                    currentQty: 0,
                    pog: 0,
                    sellIn: 0,
                    sellOut: 0,
                    totalInv: 0,
                    idleStock: 0,
                    category: item.category,
                  };
                }
                leafGroups[leafKey].lastQty += Number(item.lastQty) || 0;
                leafGroups[leafKey].currentQty += Number(item.currentQty) || 0;
                leafGroups[leafKey].sellIn += Number(item.sellIn) || 0;
                leafGroups[leafKey].sellOut += Number(item.sellOut) || 0;
                leafGroups[leafKey].totalInv += Number(item.totalInv) || 0;
                leafGroups[leafKey].pog += Number(item.pog) || 0;
                leafGroups[leafKey].idleStock += Number(item.idleStock) || 0;
              });

              let leaves = Object.entries(leafGroups).map(
                ([leafName, val]) => ({
                  name: leafName,
                  isLeaf: true,
                  ...(val as any),
                }),
              );

              if (pogSubGroupBy === "channel" || pogSubGroupBy === "kiosk") {
                const catOrder: Record<string, number> = {
                  Distributor: 1,
                  R1: 2,
                  R2: 3,
                };
                leaves.sort((a, b) => {
                  const wA = catOrder[a.category] || 99;
                  const wB = catOrder[b.category] || 99;
                  if (wA !== wB) return wA - wB;
                  return b.pog - a.pog;
                });
              } else {
                leaves.sort((a, b) => b.pog - a.pog);
              }
              finalChildrenOfSelf = leaves;
            }

            const selfNode = {
              name: node.name,
              level: node.level,
              children: finalChildrenOfSelf,
              isExpandable: finalChildrenOfSelf.length > 0,
              teamChildren: [],
              ...ownPog,
            };

            return [selfNode, ...(node.teamChildren || [])];
          } else {
            return node.teamChildren || [];
          }
        }
        return [node];
      });
    }

    const groups: Record<string, any> = {};
    filteredData.forEach((item) => {
      let key = item.hybrid || "Unknown";
      if (pogGroupBy === "area") key = item.area || "Unknown";
      else if (pogGroupBy === "category") key = item.category || "Unknown";
      else if (pogGroupBy === "crops") key = item.crops || "Unknown";

      if (!groups[key]) {
        groups[key] = {
          name: key,
          lastQty: 0,
          currentQty: 0,
          pog: 0,
          sellIn: 0,
          sellOut: 0,
          totalInv: 0,
          idleStock: 0,
          childrenMap: {},
        };
        if (pogSubGroupBy === "subordinate") {
          activePogMembers.forEach((m) => {
            groups[key].childrenMap[m] = {
              name: m,
              lastQty: 0,
              currentQty: 0,
              pog: 0,
              sellIn: 0,
              sellOut: 0,
              totalInv: 0,
              idleStock: 0,
            };
          });
        }
      }
      groups[key].lastQty += item.lastQty;
      groups[key].currentQty += item.currentQty;
      groups[key].sellIn += item.sellIn || 0;
      groups[key].sellOut += item.sellOut || 0;
      groups[key].totalInv += item.totalInv || 0;
      groups[key].pog += item.pog;
      groups[key].idleStock += item.idleStock || 0;

      // Drill down
      let subKey = "Unknown";
      if (pogSubGroupBy === "channel") subKey = item.kiosk || "Unknown Channel";
      else if (pogSubGroupBy === "hybrid") subKey = item.hybrid || "Unknown";
      else if (pogSubGroupBy === "subordinate") {
        const matched = activePogMembers.find(
          (m) => cleanForMatch(m) === cleanForMatch(item.pic),
        );
        subKey = matched || item.pic || "Unknown";
      } else if (pogSubGroupBy === "area") subKey = item.area || "Unknown";
      else if (pogSubGroupBy === "category")
        subKey = item.category || "Unknown";
      else if (pogSubGroupBy === "crops") subKey = item.crops || "Unknown";

      if (!groups[key].childrenMap[subKey]) {
        groups[key].childrenMap[subKey] = {
          name: subKey,
          lastQty: 0,
          currentQty: 0,
          pog: 0,
          sellIn: 0,
          sellOut: 0,
          totalInv: 0,
          idleStock: 0,
        };
        if (pogSubGroupBy === "channel") {
          groups[key].childrenMap[subKey].category = item.category;
        }
      }
      groups[key].childrenMap[subKey].lastQty += item.lastQty;
      groups[key].childrenMap[subKey].currentQty += item.currentQty;
      groups[key].childrenMap[subKey].sellIn += item.sellIn || 0;
      groups[key].childrenMap[subKey].sellOut += item.sellOut || 0;
      groups[key].childrenMap[subKey].totalInv += item.totalInv || 0;
      groups[key].childrenMap[subKey].pog += item.pog;
      groups[key].childrenMap[subKey].idleStock += item.idleStock || 0;
    });

    const subMembersUplines: Record<string, string | null> = {};
    const subDirectSubsMap: Record<string, string[]> = {};
    let subRootMembers: string[] = [];
    if (pogSubGroupBy === "subordinate") {
      activePogMembers.forEach((m) => {
        subMembersUplines[m] = getUplineInTeam(
          m,
          activePogMembers,
          teamUpLines,
        );
      });
      subRootMembers = activePogMembers.filter(
        (m) => subMembersUplines[m] === null,
      );
      subRootMembers.sort((a, b) =>
        compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
      );

      activePogMembers.forEach((m) => {
        const upl = subMembersUplines[m];
        if (upl !== null) {
          const uplClean = cleanForMatch(upl);
          if (!subDirectSubsMap[uplClean]) {
            subDirectSubsMap[uplClean] = [];
          }
          subDirectSubsMap[uplClean].push(m);
        }
      });

      Object.keys(subDirectSubsMap).forEach((key) => {
        subDirectSubsMap[key].sort((a, b) =>
          compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
        );
      });
    }

    return Object.values(groups)
      .map((g: any) => {
        let children = [];

        if (pogSubGroupBy === "subordinate") {
          const buildTeamPogNode = (mName: string): any => {
            const mNameClean = cleanForMatch(mName);
            const directSubs = subDirectSubsMap[mNameClean] || [];

            const directKData = g.childrenMap[mName] || {
              name: mName,
              lastQty: 0,
              currentQty: 0,
              pog: 0,
              sellIn: 0,
              sellOut: 0,
              totalInv: 0,
              idleStock: 0,
            };
            const childrenNodes = directSubs
              .map((sub) => buildTeamPogNode(sub))
              .filter(Boolean);

            const transData = { ...directKData };
            childrenNodes.forEach((child) => {
              transData.lastQty += child.lastQty || 0;
              transData.currentQty += child.currentQty || 0;
              transData.sellIn += child.sellIn || 0;
              transData.sellOut += child.sellOut || 0;
              transData.totalInv += child.totalInv || 0;
              transData.pog += child.pog || 0;
              transData.idleStock += child.idleStock || 0;
            });

            const hasSelfActivity =
              directKData.lastQty > 0 ||
              directKData.currentQty > 0 ||
              directKData.sellIn > 0 ||
              directKData.sellOut > 0 ||
              directKData.totalInv > 0 ||
              directKData.pog !== 0;
            const hasChildrenActivity = childrenNodes.length > 0;

            if (
              pogGroupBy !== "subordinate" &&
              !hasSelfActivity &&
              !hasChildrenActivity
            ) {
              return null;
            }

            return {
              name: mName,
              level: getMemberLevel(mName, teamLevels, teamPositions, userData),
              children: childrenNodes,
              isExpandable: childrenNodes.length > 0,
              teamChildren: childrenNodes,
              ...transData,
            };
          };

          children = subRootMembers
            .map((m) => buildTeamPogNode(m))
            .filter(Boolean)
            .flatMap((node) => {
              if (cleanForMatch(node.name) === cleanForMatch(userData?.name)) {
                if (userLevel <= 3) {
                  const directKData = g.childrenMap[node.name] || {
                    lastQty: 0,
                    currentQty: 0,
                    pog: 0,
                    sellIn: 0,
                    sellOut: 0,
                    totalInv: 0,
                    idleStock: 0,
                  };
                  const selfNode = {
                    name: node.name,
                    level: node.level,
                    children: [],
                    isExpandable: false,
                    teamChildren: [],
                    ...directKData,
                  };
                  return [selfNode, ...(node.teamChildren || [])];
                } else {
                  return node.teamChildren || [];
                }
              }
              return [node];
            });

          const hasGroupActivity =
            g.lastQty > 0 ||
            g.currentQty > 0 ||
            g.sellIn > 0 ||
            g.sellOut > 0 ||
            g.totalInv > 0 ||
            g.pog !== 0;
          if (children.length === 0 && !hasGroupActivity) {
            return null;
          }
        } else {
          children = Object.values(g.childrenMap);

          if (pogSubGroupBy === "channel") {
            const catOrder: Record<string, number> = {
              Distributor: 1,
              R1: 2,
              R2: 3,
            };
            children.sort((a: any, b: any) => {
              const wA = catOrder[a.category] || 99;
              const wB = catOrder[b.category] || 99;
              if (wA !== wB) return wA - wB;
              return b.pog - a.pog;
            });
          } else {
            children.sort((a: any, b: any) => b.pog - a.pog);
          }
        }

        return {
          name: g.name,
          lastQty: g.lastQty,
          currentQty: g.currentQty,
          sellIn: g.sellIn,
          sellOut: g.sellOut,
          totalInv: g.totalInv,
          idleStock: g.idleStock,
          pog: g.pog,
          isExpandable: true,
          children,
        };
      })
      .filter(Boolean)
      .sort((a, b) => b.pog - a.pog);
  }, [
    pogDataProcessed,
    pogGroupBy,
    pogSubGroupBy,
    filterBelowCrop,
    activePogMembers,
    kiosks,
    userData,
    teamSubordinates,
    teamPositions,
    teamUpLines,
    teamLevels,
    teamAreas,
    employees,
  ]);

  // Ekstrak data list crops unik untuk filter
  const availableCrops = useMemo(() => {
    const cropSet = new Set<string>();
    pogDataProcessed.forEach((item) => {
      const crop = String(item.crops || "").trim();
      if (crop) {
        cropSet.add(crop);
      }
    });
    return ["All", ...Array.from(cropSet).sort()];
  }, [pogDataProcessed]);

  // Filter team members based on selected crop filter
  const activeTeamMembers = useMemo(() => {
    if (!filterBelowCrop || filterBelowCrop === "All") return teamMembers;

    // Fallback mode: if group column info is empty or not in Google Sheet yet,
    // filter members by their matching recorded crop in workingData so they always see correct data.
    // A member is kept if they (or any of their subordinates recursively) have working data with the selected crop.
    const kiosksMapByCleanName: Record<string, any> = {};
    kiosks.forEach((k) => {
      kiosksMapByCleanName[cleanForMatch(k.name)] = k;
    });
    const teamMembersMapByCleanName: Record<string, string> = {};
    teamMembers.forEach((m) => {
      teamMembersMapByCleanName[cleanForMatch(m)] = m;
    });
    const getTeamMemberMatch = (name: string): string | undefined => {
      const clean = cleanForMatch(name);
      if (teamMembersMapByCleanName[clean])
        return teamMembersMapByCleanName[clean];
      return teamMembers.find((m) => matchNames(m, name));
    };

    const matchedMembersWithCrop = new Set<string>();
    workingData.forEach((item) => {
      const itemCrop = String(item.crops || "")
        .trim()
        .toLowerCase();
      if (checkCropMatch(itemCrop, filterBelowCrop)) {
        const kClean = cleanForMatch(item.kiosk);
        const kioskInfo = kiosksMapByCleanName[kClean] || {};
        const rawPic = normalizeName(
          String(item.user || kioskInfo.pic || "Unknown"),
        );
        const picName = getDdaOfUser(
          rawPic,
          userData?.name,
          computedTeamProfiles,
        );
        const matchedMember = getTeamMemberMatch(picName);
        if (matchedMember) {
          matchedMembersWithCrop.add(cleanForMatch(matchedMember));
        }
      }
    });

    // Propagate active status upwards
    const activeSet = new Set<string>();
    activeSet.add(cleanForMatch(userData?.name)); // Always keep self

    teamMembers.forEach((m) => {
      const mClean = cleanForMatch(m);
      if (matchedMembersWithCrop.has(mClean)) {
        activeSet.add(mClean);
        let current = m;
        for (let i = 0; i < 15; i++) {
          // Max depth protection
          const upRaw = getFromRecord<string>(teamUpLines, current);
          if (!upRaw) break;

          // Match against actual teamMembers to handle partial names
          const up = getTeamMemberMatch(upRaw as string) || (upRaw as string);

          const upClean = cleanForMatch(up);
          if (activeSet.has(upClean)) {
            // Already processed this branch upwards, can break early
            break;
          }
          activeSet.add(upClean);
          current = up;
        }
      }
    });

    return teamMembers.filter((m) => activeSet.has(cleanForMatch(m)));
  }, [
    teamMembers,
    filterBelowCrop,
    employees,
    userData,
    workingData,
    kiosks,
    teamUpLines,
  ]);

  const summaryData = useMemo(() => {
    const picToUplineMap: Record<string, string> = {};
    kiosks.forEach((k) => {
      const p = String(k.pic || "").trim();
      const u = String(k.upline || "").trim();
      if (p && u) {
        picToUplineMap[p.toLowerCase()] = normalizeName(u);
      }
    });
    picToUplineMap["listianto"] = "AGUS HERDIANTO";

    // Pre-compute maps to optimize from O(N * M) to O(N + M)
    const kiosksMapByCleanName: Record<string, any> = {};
    kiosks.forEach((k) => {
      kiosksMapByCleanName[cleanForMatch(k.name)] = k;
    });

    const teamMembersMapByCleanName: Record<string, string> = {};
    activeTeamMembers.forEach((m) => {
      teamMembersMapByCleanName[cleanForMatch(m)] = m;
    });
    const getTeamMemberMatch = (name: string): string | undefined => {
      const clean = cleanForMatch(name);
      if (teamMembersMapByCleanName[clean])
        return teamMembersMapByCleanName[clean];
      return activeTeamMembers.find((m) => matchNames(m, name));
    };

    const parseDate = (timestamp: any) => {
      if (!timestamp) return null;
      if (timestamp instanceof Date) return timestamp;
      let d = new Date(timestamp);
      if (!isNaN(d.getTime())) return d;

      if (typeof timestamp === "string") {
        const isSlash = timestamp.includes("/");
        const isDash = timestamp.includes("-");
        if (isSlash || isDash) {
          const parts = timestamp.split(/[\s/:-]+/);
          if (parts.length >= 3) {
            if (parts[0].length === 4) {
              const dStr = `${parts[0]}-${parts[1]}-${parts[2]}T${parts[3] || "00"}:${parts[4] || "00"}:${parts[5] || "00"}`;
              d = new Date(dStr);
            } else {
              let year = parts[2];
              if (year.length === 2 && !isNaN(Number(year))) {
                year = "20" + year;
              }
              const dStr = `${year}-${parts[1]}-${parts[0]}T${parts[3] || "00"}:${parts[4] || "00"}:${parts[5] || "00"}`;
              d = new Date(dStr);
            }
          }
        }
      }
      return d && !isNaN(d.getTime()) ? d : null;
    };

    const enrichedData = workingData.map((item) => {
      const d = parseDate(item.timestamp);
      const itemMonth = d ? d.getMonth() : null;
      const itemYear = d ? d.getFullYear() : null;

      const kClean = cleanForMatch(item.kiosk);
      const kioskInfo = kiosksMapByCleanName[kClean] || {};
      const rawPic = normalizeName(
        String(item.user || kioskInfo.pic || "Unknown"),
      );
      const pic = getDdaOfUser(rawPic, userData?.name, computedTeamProfiles);
      let upline = normalizeName(String(kioskInfo.upline || ""));
      if (pic.toLowerCase() === "listianto") {
        upline = "AGUS HERDIANTO";
      } else {
        const matchedMember = getTeamMemberMatch(pic);
        const foundUp = getFromRecord<string>(
          teamUpLines,
          matchedMember || pic,
        );
        if (foundUp) {
          upline = normalizeName(foundUp);
        } else if (!upline && pic !== "Unknown") {
          const foundUpline = picToUplineMap[pic.toLowerCase()];
          if (foundUpline) upline = normalizeName(foundUpline);
        }
      }
      let area = "-";
      if (item.area && String(item.area).trim() !== "") {
        area = String(item.area).trim();
      } else {
        const cleanPic = cleanForMatch(pic);
        const matchedMember = getTeamMemberMatch(pic);
        const foundArea = getFromRecord<string>(teamAreas, matchedMember || pic);
        if (foundArea) {
          area = foundArea;
        } else if (cleanPic === cleanForMatch(userData?.name)) {
          area = userData?.area || "-";
        }
      }

      const category = String(kioskInfo.category || "Uncategorized").trim();

      let cluster = "Uncategorized";
      const aging = Number(item.aging);
      if (!isNaN(aging)) {
        if (aging <= 2) cluster = "0-2";
        else if (aging <= 4) cluster = "2-4";
        else if (aging <= 6) cluster = "4-6";
        else if (aging <= 9) cluster = "6-9";
        else if (aging <= 12) cluster = "9-12";
        else cluster = ">12";
      }
      const crops =
        item.crops && String(item.crops).trim() !== ""
          ? item.crops
          : "Uncategorized Crops";
      return {
        ...item,
        pic,
        upline,
        area,
        rawArea: item.area || "-",
        category,
        cluster,
        hybrid: item.hybrid || "Unknown",
        crops,
        itemMonth,
        itemYear,
      };
    });

    enrichedSummaryDataRef.current = enrichedData;

    const teamMembersCleanSet = new Set(
      activeTeamMembers.map((m) => cleanForMatch(m)).filter(Boolean),
    );
    

    // Filter Team AND Filter Crops (Using Set.has for O(1) performance)
    const currentMonth = new Date().getMonth();
    const currentYear = new Date().getFullYear();

        const teamData = enrichedData.filter((item) => {
      const picClean = cleanForMatch(item.pic);
      let isTeamMember =
        teamMembersCleanSet.has(picClean);
        
      if (!isTeamMember && computedTeamProfiles) {
        const emp = computedTeamProfiles[picClean];
        if (emp && emp.name) {
          const empNameClean = cleanForMatch(emp.name);
          if (teamMembersCleanSet.has(empNameClean)) {
            isTeamMember = true;
          }
        }
      }

      const isCropMatch = checkCropMatch(item.crops, filterBelowCrop);
      return isTeamMember && isCropMatch;
    });

    if (summaryGroupBy === "subordinate") {
      const teamMembersUplines: Record<string, string | null> = {};
      const directSubsMap: Record<string, string[]> = {};
      activeTeamMembers.forEach((m) => {
        teamMembersUplines[m] = getUplineInTeam(
          m,
          activeTeamMembers,
          teamUpLines,
        );
      });

      const rootMembers = activeTeamMembers.filter((m) => {
        return teamMembersUplines[m] === null;
      });
      rootMembers.sort((a, b) =>
        compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
      );

      activeTeamMembers.forEach((m) => {
        const upl = teamMembersUplines[m];
        if (upl !== null) {
          const uplClean = cleanForMatch(upl);
          if (!directSubsMap[uplClean]) {
            directSubsMap[uplClean] = [];
          }
          directSubsMap[uplClean].push(m);
        }
      });

      Object.keys(directSubsMap).forEach((key) => {
        directSubsMap[key].sort((a, b) =>
          compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
        );
      });

      const buildNode = (name: string): any => {
        const nameClean = cleanForMatch(name);
        const directSubs = directSubsMap[nameClean] || [];
        const myItems = teamData.filter(
          (item) => cleanForMatch(item.pic) === nameClean,
        );
        const childrenNodes = directSubs.map((sub) => buildNode(sub));

        const ownData = {
          "0-2": 0,
          "2-4": 0,
          "4-6": 0,
          "6-9": 0,
          "9-12": 0,
          ">12": 0,
          Uncategorized: 0,
          total: 0,
        };
        myItems.forEach((item) => {
          const stock = Number(item.stock) || 0;
          if (ownData[item.cluster] !== undefined)
            ownData[item.cluster] += stock;
          ownData.total += stock;
        });

        const transData = { ...ownData };
        childrenNodes.forEach((child) => {
          Object.keys(ownData).forEach((key) => {
            if (key !== "total") {
              transData[key] += child[key] || 0;
            }
          });
          transData.total += child.total || 0;
        });

        let finalChildren = [];

        if (summarySubGroupBy === "subordinate") {
          finalChildren = [...childrenNodes];
        } else {
          const leafGroups: Record<string, any> = {};

          // Own items
          myItems.forEach((item) => {
            let leafKey = "Unknown";
            if (
              summarySubGroupBy === "channel" ||
              summarySubGroupBy === "kiosk"
            )
              leafKey = item.kiosk || "Unknown Channel";
            else if (summarySubGroupBy === "hybrid") leafKey = item.hybrid;
            else if (summarySubGroupBy === "area") leafKey = item.area;
            else if (summarySubGroupBy === "category") leafKey = item.category;
            else if (summarySubGroupBy === "crops") leafKey = item.crops;

            if (!leafGroups[leafKey]) {
              leafGroups[leafKey] = {
                "0-2": 0,
                "2-4": 0,
                "4-6": 0,
                "6-9": 0,
                "9-12": 0,
                ">12": 0,
                Uncategorized: 0,
                total: 0,
                category: item.category,
              };
            }
            const stock = Number(item.stock) || 0;
            if (leafGroups[leafKey][item.cluster] !== undefined) {
              leafGroups[leafKey][item.cluster] += stock;
            }
            leafGroups[leafKey].total += stock;
          });

          // Roll up children's leaves
          childrenNodes.forEach((child) => {
            (child.children || []).forEach((cNode: any) => {
              const leafKey = cNode.name;
              if (!leafGroups[leafKey]) {
                leafGroups[leafKey] = {
                  "0-2": 0,
                  "2-4": 0,
                  "4-6": 0,
                  "6-9": 0,
                  "9-12": 0,
                  ">12": 0,
                  Uncategorized: 0,
                  total: 0,
                  category: cNode.category,
                };
              }
              Object.keys(cNode).forEach((k) => {
                if (
                  k !== "name" &&
                  k !== "isLeaf" &&
                  k !== "category" &&
                  k !== "total" &&
                  k !== "isExpandable" &&
                  k !== "children" &&
                  k !== "level" &&
                  k !== "selectedTotal"
                ) {
                  leafGroups[leafKey][k] =
                    (leafGroups[leafKey][k] || 0) + (cNode[k] || 0);
                }
              });
              leafGroups[leafKey].total += cNode.total || 0;
            });
          });

          let leaves = Object.entries(leafGroups).map(([leafName, val]) => ({
            name: leafName,
            isLeaf: true,
            ...(val as any),
          }));

          if (
            summarySubGroupBy === "channel" ||
            summarySubGroupBy === "kiosk"
          ) {
            const catOrder: Record<string, number> = {
              Distributor: 1,
              R1: 2,
              R2: 3,
            };
            leaves.sort((a, b) => {
              const wA = catOrder[a.category] || 99;
              const wB = catOrder[b.category] || 99;
              if (wA !== wB) return wA - wB;
              return b.total - a.total;
            });
          } else {
            leaves.sort((a, b) => b.total - a.total);
          }

          finalChildren = leaves;
        }

        return {
          name,
          level: getMemberLevel(name, teamLevels, teamPositions, userData),
          children: finalChildren,
          isExpandable: finalChildren.length > 0,
          teamChildren: childrenNodes,
          ...transData,
        };
      };

      const roots = rootMembers.map((root) => buildNode(root));
      roots.sort((a, b) => b.total - a.total);
      return roots.flatMap((node) => {
        if (cleanForMatch(node.name) === cleanForMatch(userData?.name)) {
          if (userLevel <= 3) {
            const nameClean = cleanForMatch(node.name);
            const myItems = teamData.filter(
              (item) => cleanForMatch(item.pic) === nameClean,
            );

            const ownData = {
              "0-2": 0,
              "2-4": 0,
              "4-6": 0,
              "6-9": 0,
              "9-12": 0,
              ">12": 0,
              Uncategorized: 0,
              total: 0,
            };
            myItems.forEach((item) => {
              const stock = Number(item.stock) || 0;
              if (ownData[item.cluster] !== undefined)
                ownData[item.cluster] += stock;
              ownData.total += stock;
            });

            let finalChildrenOfSelf = [];
            if (summarySubGroupBy !== "subordinate") {
              const leafGroups: Record<string, any> = {};
              myItems.forEach((item) => {
                let leafKey = "Unknown";
                if (
                  summarySubGroupBy === "channel" ||
                  summarySubGroupBy === "kiosk"
                )
                  leafKey = item.kiosk || "Unknown Channel";
                else if (summarySubGroupBy === "hybrid") leafKey = item.hybrid;
                else if (summarySubGroupBy === "area") leafKey = item.area;
                else if (summarySubGroupBy === "category")
                  leafKey = item.category;
                else if (summarySubGroupBy === "crops") leafKey = item.crops;

                if (!leafGroups[leafKey]) {
                  leafGroups[leafKey] = {
                    "0-2": 0,
                    "2-4": 0,
                    "4-6": 0,
                    "6-9": 0,
                    "9-12": 0,
                    ">12": 0,
                    Uncategorized: 0,
                    total: 0,
                    category: item.category,
                  };
                }
                const stock = Number(item.stock) || 0;
                if (leafGroups[leafKey][item.cluster] !== undefined) {
                  leafGroups[leafKey][item.cluster] += stock;
                }
                leafGroups[leafKey].total += stock;
              });

              let leaves = Object.entries(leafGroups).map(
                ([leafName, val]) => ({
                  name: leafName,
                  isLeaf: true,
                  ...(val as any),
                }),
              );

              if (
                summarySubGroupBy === "channel" ||
                summarySubGroupBy === "kiosk"
              ) {
                const catOrder: Record<string, number> = {
                  Distributor: 1,
                  R1: 2,
                  R2: 3,
                };
                leaves.sort((a, b) => {
                  const wA = catOrder[a.category] || 99;
                  const wB = catOrder[b.category] || 99;
                  if (wA !== wB) return wA - wB;
                  return b.total - a.total;
                });
              } else {
                leaves.sort((a, b) => b.total - a.total);
              }
              finalChildrenOfSelf = leaves;
            }

            const selfNode = {
              name: node.name,
              level: node.level,
              children: finalChildrenOfSelf,
              isExpandable: finalChildrenOfSelf.length > 0,
              teamChildren: [],
              ...ownData,
            };

            return [selfNode, ...(node.teamChildren || [])];
          } else {
            return node.teamChildren || [];
          }
        }
        return [node];
      });
    }

    const groups: Record<string, any> = {};
    teamData.forEach((item) => {
      let key = "Unknown";
      if (summaryGroupBy === "hybrid") key = item.hybrid;
      else if (summaryGroupBy === "area") key = item.area;
      else if (summaryGroupBy === "category") key = item.category;
      else if (summaryGroupBy === "crops") key = item.crops;

      if (!groups[key]) {
        groups[key] = {
          "0-2": 0,
          "2-4": 0,
          "4-6": 0,
          "6-9": 0,
          "9-12": 0,
          ">12": 0,
          Uncategorized: 0,
          total: 0,
          childrenMap: {},
        };
        if (summarySubGroupBy === "subordinate") {
          activeTeamMembers.forEach((m) => {
            groups[key].childrenMap[m] = {
              "0-2": 0,
              "2-4": 0,
              "4-6": 0,
              "6-9": 0,
              "9-12": 0,
              ">12": 0,
              Uncategorized: 0,
              total: 0,
            };
          });
        }
      }
      const stock = Number(item.stock) || 0;
      if (groups[key][item.cluster] !== undefined) {
        groups[key][item.cluster] += stock;
      }
      groups[key].total += stock;

      // Sub-item
      let subKey = "Unknown";
      if (summarySubGroupBy === "channel")
        subKey = item.kiosk || "Unknown Channel";
      else if (summarySubGroupBy === "hybrid") subKey = item.hybrid;
      else if (summarySubGroupBy === "subordinate") {
        const matched = activeTeamMembers.find(
          (m) => cleanForMatch(m) === cleanForMatch(item.pic),
        );
        subKey = matched || item.pic;
      } else if (summarySubGroupBy === "area") subKey = item.area;
      else if (summarySubGroupBy === "category") subKey = item.category;
      else if (summarySubGroupBy === "crops") subKey = item.crops;

      if (!groups[key].childrenMap[subKey]) {
        groups[key].childrenMap[subKey] = {
          "0-2": 0,
          "2-4": 0,
          "4-6": 0,
          "6-9": 0,
          "9-12": 0,
          ">12": 0,
          Uncategorized: 0,
          total: 0,
        };
        if (summarySubGroupBy === "channel") {
          groups[key].childrenMap[subKey].category = item.category;
        }
      }
      if (groups[key].childrenMap[subKey][item.cluster] !== undefined) {
        groups[key].childrenMap[subKey][item.cluster] += stock;
      }
      groups[key].childrenMap[subKey].total += stock;
    });

    const sumMembersUplines: Record<string, string | null> = {};
    const sumDirectSubsMap: Record<string, string[]> = {};
    let sumRootMembers: string[] = [];
    if (summarySubGroupBy === "subordinate") {
      activeTeamMembers.forEach((m) => {
        sumMembersUplines[m] = getUplineInTeam(
          m,
          activeTeamMembers,
          teamUpLines,
        );
      });
      sumRootMembers = activeTeamMembers.filter(
        (m) => sumMembersUplines[m] === null,
      );
      sumRootMembers.sort((a, b) =>
        compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
      );

      activeTeamMembers.forEach((m) => {
        const upl = sumMembersUplines[m];
        if (upl !== null) {
          const uplClean = cleanForMatch(upl);
          if (!sumDirectSubsMap[uplClean]) {
            sumDirectSubsMap[uplClean] = [];
          }
          sumDirectSubsMap[uplClean].push(m);
        }
      });

      Object.keys(sumDirectSubsMap).forEach((key) => {
        sumDirectSubsMap[key].sort((a, b) =>
          compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
        );
      });
    }

    return Object.entries(groups)
      .map(([name, counts]) => {
        const countsVal = counts as any;
        let children = [];

        if (summarySubGroupBy === "subordinate") {
          const buildTeamNode = (mName: string): any => {
            const mNameClean = cleanForMatch(mName);
            const directSubs = sumDirectSubsMap[mNameClean] || [];

            const directKData = countsVal.childrenMap[mName] || {
              "0-2": 0,
              "2-4": 0,
              "4-6": 0,
              "6-9": 0,
              "9-12": 0,
              ">12": 0,
              Uncategorized: 0,
              total: 0,
            };
            const childrenNodes = directSubs.map((sub) => buildTeamNode(sub));

            const transData = { ...directKData };
            childrenNodes.forEach((child) => {
              Object.keys(directKData).forEach((k) => {
                if (k !== "total" && k !== "category") {
                  transData[k] = (transData[k] || 0) + (child[k] || 0);
                }
              });
              transData.total = (transData.total || 0) + (child.total || 0);
            });

            return {
              name: mName,
              level: getMemberLevel(mName, teamLevels, teamPositions, userData),
              children: childrenNodes,
              isExpandable: childrenNodes.length > 0,
              teamChildren: childrenNodes,
              ...transData,
            };
          };

          children = sumRootMembers
            .map((m) => buildTeamNode(m))
            .flatMap((node) => {
              if (cleanForMatch(node.name) === cleanForMatch(userData?.name)) {
                if (userLevel <= 3) {
                  const directKData = countsVal.childrenMap[node.name] || {
                    "0-2": 0,
                    "2-4": 0,
                    "4-6": 0,
                    "6-9": 0,
                    "9-12": 0,
                    ">12": 0,
                    Uncategorized: 0,
                    total: 0,
                  };
                  const selfNode = {
                    name: node.name,
                    level: node.level,
                    children: [],
                    isExpandable: false,
                    teamChildren: [],
                    ...directKData,
                  };
                  return [selfNode, ...(node.teamChildren || [])];
                } else {
                  return node.teamChildren || [];
                }
              }
              return [node];
            });
        } else {
          children = Object.entries(countsVal.childrenMap || {}).map(
            ([subKeyName, kData]) => ({
              name: subKeyName,
              ...(kData as any),
            }),
          );

          if (summarySubGroupBy === "channel") {
            const catOrder: Record<string, number> = {
              Distributor: 1,
              R1: 2,
              R2: 3,
            };
            children.sort((a, b) => {
              const wA = catOrder[a.category] || 99;
              const wB = catOrder[b.category] || 99;
              if (wA !== wB) return wA - wB;
              return b.total - a.total;
            });
          } else {
            children.sort((a, b) => b.total - a.total);
          }
        }

        const { childrenMap, ...rest } = countsVal;
        return { name, isExpandable: true, children, ...rest };
      })
      .sort((a, b) => b.total - a.total);
  }, [
    workingData,
    kiosks,
    activeTeamMembers,
    summaryGroupBy,
    summarySubGroupBy,
    userData,
    filterBelowCrop,
    teamSubordinates,
    teamAreas,
    teamUpLines,
    teamPositions,
    teamLevels,
  ]);

  const filteredSummaryData = useMemo(() => {
    if (selectedClusters.length === 0) return [];

    if (summaryGroupBy === "subordinate") {
      const filterNode = (node: any): any => {
        const parentSelectedTotal = selectedClusters.reduce(
          (sum, c) => sum + (node[c] || 0),
          0,
        );

        const filteredChildren = (node.children || [])
          .map((child: any) => {
            if (child.isLeaf) {
              const leafSelectedTotal = selectedClusters.reduce(
                (sum, c) => sum + (child[c] || 0),
                0,
              );
              if (leafSelectedTotal > 0)
                return { ...child, selectedTotal: leafSelectedTotal };
              return null;
            } else {
              return filterNode(child);
            }
          })
          .filter(Boolean);

        // ALWAYS return the teammate node (do not return null if 0 total) to support "tampilkan semua nama" from sheet employee
        return {
          ...node,
          selectedTotal: parentSelectedTotal,
          children: filteredChildren,
        };
      };

      // Do not filter(Boolean) here as we want to preserve all employee roots.
      return summaryData.map((row) => filterNode(row));
    }

    const filterSubordinateNode = (node: any): any => {
      const nodeSelectedTotal = selectedClusters.reduce(
        (sum, c) => sum + (node[c] || 0),
        0,
      );
      const filteredChildren = (node.children || [])
        .map((c: any) => filterSubordinateNode(c))
        .filter(Boolean);

      // Hanya tampilkan employee yang berkorelasi saja jika group by bukan 'subordinate' (team)
      // Karyawan dianggap berkorelasi jika memiliki stock > 0 atau salah satu bawahannya memiliki stock > 0
      if (
        summaryGroupBy !== "subordinate" &&
        nodeSelectedTotal === 0 &&
        filteredChildren.length === 0
      ) {
        return null;
      }

      return {
        ...node,
        selectedTotal: nodeSelectedTotal,
        children: filteredChildren,
      };
    };

    return summaryData
      .map((row) => {
        if (row.isExpandable) {
          if (summarySubGroupBy === "subordinate") {
            const filteredChildren = row.children
              .map((child) => filterSubordinateNode(child))
              .filter(Boolean);
            const parentSelectedTotal = selectedClusters.reduce(
              (sum, c) => sum + (row[c] || 0),
              0,
            );
            if (filteredChildren.length === 0 && parentSelectedTotal === 0)
              return null;
            return {
              ...row,
              children: filteredChildren,
              selectedTotal: parentSelectedTotal,
            };
          }

          const filteredChildren = row.children
            .map((child) => {
              const childSelectedTotal = selectedClusters.reduce(
                (sum, c) => sum + (child[c] || 0),
                0,
              );
              // If sub-group is Team/Subordinate, ALWAYS show them even with 0 total
              if (
                summarySubGroupBy === "subordinate" ||
                childSelectedTotal > 0
              ) {
                return { ...child, selectedTotal: childSelectedTotal };
              }
              return null;
            })
            .filter(Boolean);

          const parentSelectedTotal = selectedClusters.reduce(
            (sum, c) => sum + (row[c] || 0),
            0,
          );
          // If sub-group is Team/Subordinate, ALWAYS show the parent group even with 0 total
          if (
            summarySubGroupBy === "subordinate" ||
            filteredChildren.length > 0 ||
            parentSelectedTotal > 0
          ) {
            return {
              ...row,
              children: filteredChildren,
              selectedTotal: parentSelectedTotal,
            };
          }
          return null;
        } else {
          const rowSelectedTotal = selectedClusters.reduce(
            (sum, c) => sum + (row[c] || 0),
            0,
          );
          if (rowSelectedTotal > 0)
            return { ...row, selectedTotal: rowSelectedTotal };
          return null;
        }
      })
      .filter(Boolean);
  }, [summaryData, selectedClusters, summaryGroupBy, summarySubGroupBy]);

  const totalSummary = useMemo(() => {
    const totals = {
      "0-2": 0,
      "2-4": 0,
      "4-6": 0,
      "6-9": 0,
      "9-12": 0,
      ">12": 0,
      Uncategorized: 0,
      selectedTotal: 0,
    };
    summaryData.forEach((row) => {
      ALL_CLUSTER_KEYS.forEach((c) => {
        totals[c] += row[c] || 0;
      });
      totals.selectedTotal += selectedClusters.reduce(
        (sum, c) => sum + (row[c] || 0),
        0,
      );
    });
    return totals;
  }, [summaryData, selectedClusters, ALL_CLUSTER_KEYS]);

  const overviewUseMt = useMemo(() => {
    const maxVal = Math.max(
      overviewStats.totalOpeningStock || 0,
      overviewStats.totalCurrentStock || 0,
      overviewStats.totalSellIn || 0,
      overviewStats.totalIdleStock || 0,
      overviewStats.totalSellOut || 0,
      totalSummary.selectedTotal || 0
    );
    return maxVal >= 1000;
  }, [overviewStats, totalSummary]);

  const handleDownloadSummaryExcel = () => {
    const teamMembersCleanSet = new Set(
      activeTeamMembers.map((m) => cleanForMatch(m)).filter(Boolean),
    );
    
    const currentMonth = new Date().getMonth();
    const currentYear = new Date().getFullYear();

    // Filter enriched data to match team members and active crop + month + chosen aging clusters
    const displayedRawItems = (enrichedSummaryDataRef.current || []).filter(
      (item) => {
        const picClean = cleanForMatch(item.pic);
        const isTeamMember =
          teamMembersCleanSet.has(picClean);
        const isCropMatch = checkCropMatch(item.crops, filterBelowCrop);
        const isClusterMatch = selectedClusters.includes(item.cluster);

        return isTeamMember && isCropMatch && isClusterMatch;
      },
    );

    if (displayedRawItems.length === 0) {
      alert("Tidak ada data untuk di-download.");
      return;
    }

    // Map each item exactly to match the custom requested headers in Indonesian and English
    const formattedRows = displayedRawItems.map((item) => {
      return {
        tgl: item.timestamp || "",
        province:
          item.province ||
          getFromRecord<string>(teamProvinces, item.pic) ||
          "-",
        crops: item.crops || "",
        checker: item.user || item.pic || "",
        channel: item.kiosk || "",
        category: item.category || "",
        hybrids: item.hybrid || "",
        "lot no": item.lot || "",
        qty: Number(item.stock) || 0,
        "usia stock": item.aging !== undefined ? String(item.aging) : "",
        "cluster aging": item.cluster || "",
        "shipping date":
          item.drDate ||
          item.shipping_date ||
          item.shippingDate ||
          item.dr_date ||
          "",
        "exp date": item.expired || "",
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(formattedRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Working Data Sheet");

    // Pre-configure column widths for pristine visual alignments to match the new headers
    const colWidths = [
      { wch: 22 }, // tgl
      { wch: 18 }, // province
      { wch: 18 }, // crops
      { wch: 25 }, // checker
      { wch: 30 }, // channel
      { wch: 18 }, // category
      { wch: 28 }, // hybrids
      { wch: 18 }, // lot no
      { wch: 12 }, // qty
      { wch: 18 }, // usia stock
      { wch: 15 }, // cluster aging
      { wch: 20 }, // shipping date
      { wch: 22 }, // exp date
    ];
    worksheet["!cols"] = colWidths;

    XLSX.writeFile(
      workbook,
      `Stock_RADAR_AWBA_${new Date().toISOString().split("T")[0]}.xlsx`,
    );
  };

  const totalPog = useMemo(() => {
    const total = {
      lastQty: 0,
      currentQty: 0,
      sellIn: 0,
      sellOut: 0,
      totalInv: 0,
      pog: 0,
      idleStock: 0,
    };
    aggregatedPogData.forEach((row) => {
      total.lastQty += row.lastQty || 0;
      total.currentQty += row.currentQty || 0;
      total.sellIn += row.sellIn || 0;
      total.sellOut += row.sellOut || 0;
      total.totalInv += row.totalInv || 0;
      total.pog += row.pog || 0;
      total.idleStock += row.idleStock || 0;
    });
    return total;
  }, [aggregatedPogData]);

  const totalSummaryPog = useMemo(() => {
    const teamMembersCleanSet = new Set(
      activeTeamMembers.map((m) => cleanForMatch(m)).filter(Boolean),
    );
    

    const filteredData = pogDataProcessed.filter((item) => {
      const picClean = cleanForMatch(item.pic);
      const isTeamMember =
        teamMembersCleanSet.has(picClean);
      const isCropMatch = checkCropMatch(item.crops, filterBelowCrop);
      return isTeamMember && isCropMatch;
    });

    const total = {
      lastQty: 0,
      currentQty: 0,
      sellIn: 0,
      sellOut: 0,
      totalInv: 0,
      pog: 0,
      idleStock: 0,
    };
    filteredData.forEach((row) => {
      total.lastQty += row.lastQty || 0;
      total.currentQty += row.currentQty || 0;
      total.sellIn += row.sellIn || 0;
      total.sellOut += row.sellOut || 0;
      total.totalInv += row.totalInv || 0;
      total.pog += row.pog || 0;
      total.idleStock += row.idleStock || 0;
    });
    return total;
  }, [pogDataProcessed, activeTeamMembers, filterBelowCrop]);

  const totalTeamStats = useMemo(() => {
    let total = 0;
    let visited = 0;
    teamStats.forEach((s) => {
      total += s.total || 0;
      visited += s.totalVisited || 0;
    });
    const percentage =
      total > 0 ? Math.min(100, Math.round((visited / total) * 100)) : 0;
    return { visited, total, percentage };
  }, [teamStats]);

  const activeVisitStats = useMemo(() => {
    const isAll =
      !mappingPic ||
      cleanForMatch(mappingPic) === "allteam" ||
      cleanForMatch(mappingPic) === "all_team";
    if (isAll) {
      return totalTeamStats;
    }
    const picStat = getStatsForPic(mappingPic);
    return {
      visited: picStat.visited,
      total: picStat.total,
      percentage: picStat.percentage,
    };
  }, [mappingPic, totalTeamStats, teamStats]);

  const grandTotalStats = useMemo(() => {
    const teamMembersCleanSet = new Set(
      activeTeamMembers.map((m) => cleanForMatch(m)).filter(Boolean),
    );
    

    const displayedRawItems = (enrichedSummaryDataRef.current || []).filter(
      (item) => {
        const picClean = cleanForMatch(item.pic);
        const isTeamMember =
          teamMembersCleanSet.has(picClean);
        const isCropMatch = checkCropMatch(item.crops, filterBelowCrop);
        const isClusterMatch = selectedClusters.includes(item.cluster);

        return isTeamMember && isCropMatch && isClusterMatch;
      },
    );

    const groupData: Record<
      string,
      { total: number; clusters: Record<string, number> }
    > = {};
    displayedRawItems.forEach((item) => {
      let g = "Unknown";
      if (grandTotalViewBy === "hybrid") {
        g = item.hybrid || "Unknown";
      } else if (grandTotalViewBy === "area") {
        const itemProv = item.province || getFromRecord<string>(teamProvinces, item.pic) || (item.area && item.area !== "-" ? item.area : "") || (item.rawArea && item.rawArea !== "-" ? item.rawArea : "");
        g = (itemProv && itemProv !== "-") ? itemProv : "Unknown";
      }
      
      if (groupData[g] === undefined) {
        groupData[g] = { total: 0, clusters: {} };
      }
      const qty = Number(item.stock) || 0;
      groupData[g].total += qty;

      const cluster = item.cluster;
      if (groupData[g].clusters[cluster] === undefined) {
        groupData[g].clusters[cluster] = 0;
      }
      groupData[g].clusters[cluster] += qty;
    });

    const sortedGroups = Object.entries(groupData)
      .sort((a, b) => b[1].total - a[1].total) // Sort descending by total kg
      .map(([name, data]) => ({
        name,
        total: data.total,
        clusters: data.clusters,
      }));

    return sortedGroups;
  }, [summaryData, activeTeamMembers, filterBelowCrop, selectedClusters, grandTotalViewBy, teamProvinces]);

  const categoryFillingStats = useMemo(() => {
    let total = 0;
    let filled = 0;

    mappedChannelsByPic.forEach((k) => {
      total++;
      const cat = String(k.category || "").trim();
      const cleanCat = cat.toLowerCase();
      if (
        cat !== "" &&
        cat !== "-" &&
        cleanCat !== "uncategorized" &&
        cleanCat !== "n/a" &&
        cleanCat !== "unknown"
      ) {
        filled++;
      }
    });

    const percentage =
      total > 0 ? Math.min(100, Math.round((filled / total) * 100)) : 0;
    return { filled, total, percentage };
  }, [mappedChannelsByPic]);

  const categoryDistributionStats = useMemo(() => {
    const counts: Record<string, number> = {};
    let total = 0;

    mappedChannelsByPic.forEach((k) => {
      total++;
      let cat = String(k.category || "").trim();
      if (cat === "" || cat === "-") {
        cat = "Uncategorized";
      }

      const lowerCat = cat.toLowerCase();
      if (
        lowerCat === "uncategorized" ||
        lowerCat === "n/a" ||
        lowerCat === "unknown"
      ) {
        cat = "Uncategorized";
      } else if (lowerCat === "r1") {
        cat = "R1";
      } else if (lowerCat === "r2") {
        cat = "R2";
      } else if (lowerCat === "distributor") {
        cat = "Distributor";
      } else {
        cat = cat.charAt(0).toUpperCase() + cat.slice(1);
      }

      counts[cat] = (counts[cat] || 0) + 1;
    });

    const data = Object.entries(counts)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => {
        const order = ["distributor", "r1", "r2"];
        const idxA = order.indexOf(a.name.toLowerCase());
        const idxB = order.indexOf(b.name.toLowerCase());

        if (idxA !== -1 && idxB !== -1) {
          return idxA - idxB;
        }
        if (idxA !== -1) return -1;
        if (idxB !== -1) return 1;

        return b.value - a.value;
      });

    return { data, total };
  }, [mappedChannelsByPic]);

  const categoryChartSegments = useMemo(() => {
    const { data, total } = categoryDistributionStats;
    if (total === 0 || data.length === 0) return [];

    let accumulatedPercent = 0;
    const circumference = 2 * Math.PI * 95; // 596.9026

    const getCategoryColor = (name: string, idx: number) => {
      const lower = name.toLowerCase();
      if (lower === "gold") return "#f59e0b"; // Gold Amber-500
      if (lower === "silver") return "#cbd5e1"; // Silver Slate-300
      if (lower === "bronze") return "#b45309"; // Bronze Amber-700

      // Primary partner categories
      if (lower === "distributor") return "rgba(255, 255, 255, 0.35)"; // White with lower opacity
      if (lower === "r1") return "#ffffff"; // White
      if (lower === "r2") return "#2563eb"; // Blue cluster color (blue-600)

      if (lower === "uncategorized" || lower === "unknown" || lower === "-")
        return "#94a3b8"; // Neutral Slate

      const COLORS = [
        "rgba(255, 255, 255, 0.35)",
        "#ffffff",
        "#2563eb",
        "#f59e0b",
        "#fb923c",
        "#f43f5e",
        "#2dd4bf",
        "#94a3b8",
      ];
      return COLORS[idx % COLORS.length];
    };

    return data.map((item, idx) => {
      const percentage = total > 0 ? (item.value / total) * 100 : 0;
      const strokeLength = (percentage / 100) * circumference;
      const strokeOffset =
        circumference - (accumulatedPercent / 100) * circumference;
      accumulatedPercent += percentage;
      return {
        ...item,
        percentage,
        strokeDasharray: `${strokeLength} ${circumference - strokeLength}`,
        strokeDashoffset: strokeOffset,
        color: getCategoryColor(item.name, idx),
      };
    });
  }, [categoryDistributionStats]);

  const renderRecursiveSummaryRow = (row: any, depth = 0): React.ReactNode => {
    const hasChildren = row.children && row.children.length > 0;
    const isExpanded = !!expandedRows[row.name];

    // Determine level badge style
    const levelVal = row.level;
    const levelLabel = levelVal !== undefined ? `Level ${levelVal}` : "";
    const isZeroTotal = row.selectedTotal === 0;

    return (
      <div
        key={row.name}
        className={`overflow-hidden transition-all ${isZeroTotal ? "bg-red-50/40 border-l-[3px] border-red-300" : ""} ${depth === 0 ? "md:bg-white md:rounded-[32px] md:shadow-[0_4px_24px_rgba(24,26,44,0.08)] md:border md:border-[#edecff]" : ""}`}
      >
        <div
          className={`flex justify-between items-center px-5 py-4 pb-2 transition-colors duration-150 animate-in fade-in slide-in-from-left-2 duration-200 ${isZeroTotal ? "hover:bg-red-100/40" : "hover:bg-slate-50/60"}`}
          style={{ paddingLeft: `${Math.max(20, depth * 20)}px` }}
        >
          <div className="flex flex-col min-w-0 flex-1 pr-2">
            <span
              className="font-semibold text-xs md:text-sm text-[#181a2c] uppercase flex items-center gap-1.5 cursor-pointer select-none"
              onClick={() => {
                if (hasChildren) {
                  toggleRow(row.name);
                }
              }}
            >
              {hasChildren && (
                <span
                  className={`material-symbols-outlined text-[20px] shrink-0 ${isZeroTotal ? "text-red-500" : "text-primary"}`}
                >
                  {isExpanded ? "keyboard_arrow_down" : "keyboard_arrow_right"}
                </span>
              )}
              <span className="truncate">
                {renderMaybeChannelName(row.name)}
              </span>
              {isZeroTotal && (
                <span className="text-[7.5px] font-extrabold bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full shrink-0 uppercase tracking-wider">
                  No Activity
                </span>
              )}
              {row.isLeaf &&
                summarySubGroupBy === "channel" &&
                row.category && (
                  <span className="text-[8px] bg-primary/10 text-primary px-1.5 py-0.5 rounded-full font-bold">
                    {row.category}
                  </span>
                )}
            </span>
            {!row.isLeaf && getStatsForPic(row.name) && (
              <div className="flex gap-1.5 mt-1 ml-6 text-[8.5px] uppercase tracking-wide font-bold text-[#8E94B7] flex-wrap">
                <span className="bg-[#f4f2ff] px-2 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                  V:{" "}
                  <span className="text-primary">
                    {getStatsForPic(row.name)?.visited}
                  </span>
                </span>
                <span className="bg-[#f4f2ff] px-2 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                  C:{" "}
                  <span className="text-[#181a2c]">
                    {getStatsForPic(row.name)?.total}
                  </span>
                </span>
                <span className="bg-[#f4f2ff] px-2 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                  %:{" "}
                  <span className="text-emerald-600">
                    {getStatsForPic(row.name)?.percentage}%
                  </span>
                </span>
              </div>
            )}
          </div>
          <div className="flex flex-col items-end shrink-0">
            <span
              className={`font-bold text-sm ${isZeroTotal ? "text-red-600" : "text-primary"}`}
            >
              {formatOverviewVal(row.selectedTotal, overviewUseMt).valueStr}
            </span>
            <span className="text-[8px] text-[#8E94B7] uppercase tracking-widest font-bold">
              Total {formatOverviewVal(row.selectedTotal, overviewUseMt).unit}
            </span>
          </div>
        </div>

        {/* Value Clusters in primary container */}
        <div
          className={`mx-5 mb-3 mt-2 flex divide-x rounded-[14px] overflow-hidden ${isZeroTotal ? "divide-red-200 bg-red-300 shadow-[0_12px_32px_rgba(239,68,68,0.18)]" : "divide-white/20 bg-primary shadow-[0_12px_32px_rgba(21,75,226,0.35)]"}`}
          style={{ marginLeft: `${Math.max(20, depth * 20)}px` }}
        >
          {selectedClusters.map((clusterKey) => {
            if (
              clusterKey === "Uncategorized" &&
              (!row[clusterKey] || row[clusterKey] === 0)
            )
              return null;
            const clusterConfig = CLUSTER_CONFIG.find(
              (c) => c.key === clusterKey,
            );
            return (
              <div
                key={clusterKey}
                className={`flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center transition-colors ${isZeroTotal ? "hover:bg-white/15" : "hover:bg-white/10"}`}
              >
                <span
                  className={`text-[8px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isZeroTotal ? "text-white/80" : "text-white/85"}`}
                >
                  {clusterConfig?.label || clusterKey}
                </span>
                <span className="font-semibold text-[10.5px] truncate w-full text-white">
                  {formatOverviewVal(row[clusterKey], overviewUseMt).valueStr}
                </span>
              </div>
            );
          })}
        </div>

        {isExpanded && hasChildren && (
          <div className="pb-2 border-t border-[#edecff] bg-slate-50/15">
            {row.children.map((child: any) =>
              renderRecursiveSummaryRow(child, depth + 1),
            )}
          </div>
        )}
      </div>
    );
  };

  const renderRecursiveSubordinate = (
    child: any,
    depth = 0,
  ): React.ReactNode => {
    const isChildZeroTeam = child.selectedTotal === 0;
    const hasSubChildren = child.children && child.children.length > 0;
    const isSubExpanded = !!expandedRows[child.name];

    return (
      <div
        key={child.name}
        className="flex flex-col w-full animate-in fade-in slide-in-from-left-2 duration-200"
        style={{ paddingLeft: depth > 0 ? "16px" : "0px" }}
      >
        <div
          className={`flex flex-col p-3.5 rounded-[18px] transition-all duration-200 mb-2 ${
            isChildZeroTeam
              ? "bg-red-50/70 border border-red-200/60 shadow-[0_10px_28px_rgba(239,68,68,0.12)]"
              : "bg-[#fbfaff] shadow-[0_10px_28px_rgba(21,75,226,0.18)]"
          }`}
        >
          <div className="flex justify-between items-center mb-2 px-1 flex-wrap gap-2">
            <div className="flex flex-col min-w-0">
              <div className="flex items-center gap-1.5">
                {hasSubChildren && (
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      toggleRow(child.name);
                    }}
                    className="focus:outline-none flex items-center justify-center p-0.5 hover:bg-slate-200/50 rounded-full"
                  >
                    <span className="material-symbols-outlined text-primary text-[16px]">
                      {isSubExpanded
                        ? "keyboard_arrow_down"
                        : "keyboard_arrow_right"}
                    </span>
                  </button>
                )}
                <span className="font-bold text-[11px] text-[#181a2c] uppercase flex items-center gap-1.5 flex-wrap">
                  <span className="truncate">
                    {renderMaybeChannelName(child.name)}
                  </span>
                  {isChildZeroTeam && (
                    <span className="text-[7.5px] bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full shrink-0 uppercase tracking-wider animate-pulse font-bold">
                      No Activity
                    </span>
                  )}
                </span>
              </div>
              {getStatsForPic(child.name) && (
                <div className="flex gap-1.2 mt-1 ml-5 text-[8px] uppercase tracking-wide font-bold text-[#8E94B7] flex-wrap">
                  <span className="bg-[#f4f2ff] px-1.5 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                    V:{" "}
                    <span className="text-primary">
                      {getStatsForPic(child.name)?.visited}
                    </span>
                  </span>
                  <span className="bg-[#f4f2ff] px-1.5 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                    C:{" "}
                    <span className="text-[#181a2c]">
                      {getStatsForPic(child.name)?.total}
                    </span>
                  </span>
                  <span className="bg-[#f4f2ff] px-1.5 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                    %:{" "}
                    <span className="text-emerald-600">
                      {getStatsForPic(child.name)?.percentage}%
                    </span>
                  </span>
                </div>
              )}
            </div>
            <span
              className={`font-bold text-[11.5px] shrink-0 ${isChildZeroTeam ? "text-red-600" : "text-[#181a2c]"}`}
            >
              {formatOverviewVal(child.selectedTotal, overviewUseMt).valueStr}{" "}
              <span className="text-[8.5px] text-[#8E94B7]">{formatOverviewVal(child.selectedTotal, overviewUseMt).unit}</span>
            </span>
          </div>
          <div
            className={`flex w-full divide-x rounded-[14px] overflow-hidden ${
              isChildZeroTeam
                ? "divide-red-200 bg-red-100/40"
                : "divide-primary/10 bg-primary/8"
            }`}
          >
            {selectedClusters.map((clusterKey) => {
              if (
                clusterKey === "Uncategorized" &&
                (!child[clusterKey] || child[clusterKey] === 0)
              )
                return null;
              const clusterConfig = CLUSTER_CONFIG.find(
                (c) => c.key === clusterKey,
              );
              return (
                <div
                  key={clusterKey}
                  className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                    isChildZeroTeam
                      ? "hover:bg-red-100/30"
                      : "hover:bg-primary/5"
                  }`}
                >
                  <span
                    className={`text-[7px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isChildZeroTeam ? "text-red-700/80" : "text-[#8E94B7]"}`}
                  >
                    {clusterConfig?.label || clusterKey}
                  </span>
                  <span
                    className={`font-semibold text-[9.5px] truncate w-full ${isChildZeroTeam ? "text-red-600" : "text-[#181a2c]"}`}
                  >
                    {formatOverviewVal(child[clusterKey], overviewUseMt).valueStr}
                  </span>
                </div>
              );
            })}
          </div>
        </div>

        {isSubExpanded && hasSubChildren && (
          <div className="flex flex-col border-l border-[#edecff] ml-3 pl-1 gap-1 mb-2">
            {child.children.map((subChild: any) =>
              renderRecursiveSubordinate(subChild, depth + 1),
            )}
          </div>
        )}
      </div>
    );
  };

  const renderRecursivePogSubordinate = (
    child: any,
    depth = 0,
  ): React.ReactNode => {
    const isChildZeroTeam =
      child.lastQty === 0 &&
      child.sellIn === 0 &&
      child.sellOut === 0 &&
      child.totalInv === 0 &&
      child.currentQty === 0;
    const hasSubChildren = child.children && child.children.length > 0;
    const isSubExpanded = !!pogExpandedRows[child.name];

    return (
      <div
        key={child.name}
        className="flex flex-col w-full animate-in fade-in slide-in-from-left-2 duration-200"
        style={{ paddingLeft: depth > 0 ? "16px" : "0px" }}
      >
        <div
          className={`flex flex-col p-3.5 rounded-[18px] transition-all duration-200 mb-2 ${
            isChildZeroTeam
              ? "bg-red-50/70 border border-red-200/60 shadow-[0_10px_28px_rgba(239,68,68,0.12)]"
              : "bg-[#fbfaff] shadow-[0_10px_28px_rgba(21,75,226,0.18)]"
          }`}
        >
          <div className="flex justify-between items-center mb-2 px-1 flex-wrap gap-2">
            <div className="flex flex-col min-w-0">
              <div className="flex items-center gap-1.5">
                {hasSubChildren && (
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      togglePogRow(child.name);
                    }}
                    className="focus:outline-none flex items-center justify-center p-0.5 hover:bg-slate-200/50 rounded-full"
                  >
                    <span className="material-symbols-outlined text-primary text-[16px]">
                      {isSubExpanded
                        ? "keyboard_arrow_down"
                        : "keyboard_arrow_right"}
                    </span>
                  </button>
                )}
                <span className="font-bold text-[11px] text-[#181a2c] uppercase flex items-center gap-1.5 flex-wrap">
                  <span className="truncate">
                    {renderMaybeChannelName(child.name)}
                  </span>
                  {isChildZeroTeam && (
                    <span className="text-[7.5px] bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full font-bold uppercase tracking-wider">
                      No Activity
                    </span>
                  )}
                </span>
              </div>
              {getStatsForPic(child.name) && (
                <div className="flex gap-1.2 mt-1 ml-5 text-[8px] uppercase tracking-wide font-bold text-[#8E94B7] flex-wrap">
                  <span className="bg-[#f4f2ff] px-1.5 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                    V:{" "}
                    <span className="text-primary">
                      {getStatsForPic(child.name)?.visited}
                    </span>
                  </span>
                  <span className="bg-[#f4f2ff] px-1.5 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                    C:{" "}
                    <span className="text-[#181a2c]">
                      {getStatsForPic(child.name)?.total}
                    </span>
                  </span>
                  <span className="bg-[#f4f2ff] px-1.5 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                    %:{" "}
                    <span className="text-emerald-600">
                      {getStatsForPic(child.name)?.percentage}%
                    </span>
                  </span>
                </div>
              )}
            </div>
            <span
              className={`font-bold text-[11.5px] shrink-0 ${isChildZeroTeam ? "text-red-600" : "text-primary"}`}
            >
              {formatOverviewVal(child.pog, overviewUseMt).valueStr}{" "}
              <span className="text-[8.5px] text-[#8E94B7]">POG ({formatOverviewVal(child.pog, overviewUseMt).unit})</span>
            </span>
          </div>
          <div className="flex flex-row w-full gap-1.5 md:gap-2">
            {/* Table 1: Opening Inv & End of Inv */}
            <div
              className={`flex-[2] flex divide-x rounded-[14px] overflow-hidden ${
                isChildZeroTeam
                  ? "divide-red-200 bg-red-100/40"
                  : "divide-primary/10 bg-primary/5"
              }`}
            >
              <div
                className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                  isChildZeroTeam ? "hover:bg-red-200/30" : "hover:bg-primary/5"
                }`}
              >
                <span
                  className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                    isChildZeroTeam ? "text-red-700/60" : "text-[#8E94B7]"
                  }`}
                >
                  Opening Inv
                </span>
                <span
                  className={`font-black text-[10px] truncate w-full ${
                    isChildZeroTeam ? "text-red-700" : "text-[#181a2c]"
                  }`}
                >
                  {formatOverviewVal(child.lastQty, overviewUseMt).valueStr}
                </span>
              </div>
              <div
                className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                  isChildZeroTeam ? "hover:bg-red-200/30" : "hover:bg-primary/5"
                }`}
              >
                <span
                  className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                    isChildZeroTeam ? "text-red-700/60" : "text-[#1d4ed8]/75"
                  }`}
                >
                  End of Inv
                </span>
                <span
                  className={`font-black text-[10px] truncate w-full ${
                    isChildZeroTeam ? "text-red-700" : "text-[#1d4ed8]"
                  }`}
                >
                  {formatOverviewVal(child.currentQty, overviewUseMt).valueStr}
                </span>
              </div>
            </div>

            {/* Table 2: Stock in, idle stock, POG */}
            <div
              className={`flex-[3] flex divide-x rounded-[14px] overflow-hidden ${
                isChildZeroTeam
                  ? "divide-red-200 bg-red-200/45"
                  : "divide-primary/10 bg-primary/10 border border-primary/10"
              }`}
            >
              <div
                className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                  isChildZeroTeam
                    ? "hover:bg-red-300/30"
                    : "hover:bg-primary/15"
                }`}
              >
                <span
                  className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                    isChildZeroTeam ? "text-red-800/70" : "text-[#154be2]/80"
                  }`}
                >
                  Stock in
                </span>
                <span
                  className={`font-black text-[10px] truncate w-full ${
                    isChildZeroTeam ? "text-red-800" : "text-[#154be2]"
                  }`}
                >
                  {formatOverviewVal(child.sellIn, overviewUseMt).valueStr}
                </span>
              </div>
              <div
                className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                  isChildZeroTeam
                    ? "hover:bg-red-300/30"
                    : "hover:bg-amber-100/60"
                }`}
              >
                <span
                  className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                    isChildZeroTeam ? "text-red-800/70" : "text-amber-800"
                  }`}
                >
                  idle stock
                </span>
                <span
                  className={`font-black text-[10px] truncate w-full ${
                    isChildZeroTeam ? "text-red-800" : "text-amber-700"
                  }`}
                >
                  {formatOverviewVal(child.idleStock, overviewUseMt).valueStr}
                </span>
              </div>
              <div
                className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                  isChildZeroTeam
                    ? "hover:bg-red-300/30"
                    : "hover:bg-emerald-100/60"
                }`}
              >
                <span
                  className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                    isChildZeroTeam ? "text-red-800/70" : "text-emerald-800"
                  }`}
                >
                  POG
                </span>
                <span
                  className={`font-black text-[10px] truncate w-full ${
                    isChildZeroTeam
                      ? "text-red-800"
                      : "text-emerald-700 font-extrabold"
                  }`}
                >
                  {formatOverviewVal(child.pog, overviewUseMt).valueStr}
                </span>
              </div>
            </div>
          </div>
        </div>

        {isSubExpanded && hasSubChildren && (
          <div className="flex flex-col border-l border-[#edecff] ml-3 pl-1 gap-1 mb-2">
            {child.children.map((subChild: any) =>
              renderRecursivePogSubordinate(subChild, depth + 1),
            )}
          </div>
        )}
      </div>
    );
  };

  const renderRecursivePogRow = (row: any, depth = 0): React.ReactNode => {
    const hasChildren = row.children && row.children.length > 0;
    const isExpanded = !!pogExpandedRows[row.name];

    // Determine level badge style
    const levelVal = row.level;
    const levelLabel = levelVal !== undefined ? `Level ${levelVal}` : "";
    const isZeroPogActivity =
      row.lastQty === 0 &&
      row.sellIn === 0 &&
      row.sellOut === 0 &&
      row.totalInv === 0 &&
      row.currentQty === 0;

    return (
      <div
        key={row.name}
        className={`overflow-hidden transition-all ${isZeroPogActivity ? "bg-red-50/40 border-l-[3px] border-red-300" : ""} ${depth === 0 ? "md:bg-white md:rounded-[32px] md:shadow-[0_4px_24px_rgba(24,26,44,0.08)] md:border md:border-[#edecff]" : ""}`}
      >
        <div
          className={`flex justify-between items-center px-5 py-4 pb-2 transition-colors duration-150 animate-in fade-in slide-in-from-right-2 duration-200 ${isZeroPogActivity ? "hover:bg-red-100/40" : "hover:bg-slate-50/60"}`}
          style={{ paddingLeft: `${Math.max(20, depth * 20)}px` }}
        >
          <div className="flex flex-col min-w-0 flex-1 pr-2">
            <span
              className="font-semibold text-xs md:text-sm text-[#181a2c] uppercase flex items-center gap-1.5 cursor-pointer select-none"
              onClick={() => {
                if (hasChildren) {
                  togglePogRow(row.name);
                }
              }}
            >
              {hasChildren && (
                <span
                  className={`material-symbols-outlined text-[20px] shrink-0 ${isZeroPogActivity ? "text-red-500" : "text-primary"}`}
                >
                  {isExpanded ? "keyboard_arrow_down" : "keyboard_arrow_right"}
                </span>
              )}
              <span className="truncate">
                {renderMaybeChannelName(row.name)}
              </span>
              {isZeroPogActivity && (
                <span className="text-[7.5px] font-extrabold bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full shrink-0 uppercase tracking-wider">
                  No Activity
                </span>
              )}
              {row.isLeaf && pogSubGroupBy === "channel" && row.category && (
                <span className="text-[8px] bg-primary/10 text-primary px-1.5 py-0.5 rounded-full font-bold">
                  {row.category}
                </span>
              )}
            </span>
            {!row.isLeaf && getStatsForPic(row.name) && (
              <div className="flex gap-1.5 mt-1 ml-6 text-[8.5px] uppercase tracking-wide font-bold text-[#8E94B7] flex-wrap">
                <span className="bg-[#f4f2ff] px-2 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                  V:{" "}
                  <span className="text-primary">
                    {getStatsForPic(row.name)?.visited}
                  </span>
                </span>
                <span className="bg-[#f4f2ff] px-2 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                  C:{" "}
                  <span className="text-[#181a2c]">
                    {getStatsForPic(row.name)?.total}
                  </span>
                </span>
                <span className="bg-[#f4f2ff] px-2 py-0.5 rounded-full border border-[#edecff] flex gap-1 text-emerald-600">
                  %: <span>{getStatsForPic(row.name)?.percentage}%</span>
                </span>
              </div>
            )}
          </div>
          <div className="flex flex-col items-end shrink-0">
            <span
              className={`font-bold text-sm ${isZeroPogActivity ? "text-red-600" : "text-primary"}`}
            >
              {formatOverviewVal(row.pog, overviewUseMt).valueStr}
            </span>
            <span className="text-[8px] text-[#8E94B7] uppercase tracking-widest font-bold">
              POG ({formatOverviewVal(row.pog, overviewUseMt).unit})
            </span>
          </div>
        </div>

        {/* Dynamic primary row columns */}
        <div
          className="mx-5 mb-3 mt-2 flex flex-row gap-1.5 md:gap-2"
          style={{ marginLeft: `${Math.max(20, depth * 20)}px` }}
        >
          {/* Table 1: Opening Inv & End of Inv */}
          <div
            className={`flex-[2] flex divide-x rounded-[14px] overflow-hidden ${
              isZeroPogActivity
                ? "divide-red-200 bg-red-300 shadow-[0_12px_32px_rgba(239,68,68,0.18)]"
                : "divide-white/20 bg-primary/95 shadow-[0_12px_32px_rgba(21,75,226,0.25)]"
            }`}
          >
            <div
              className={`flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center transition-colors ${isZeroPogActivity ? "hover:bg-white/15" : "hover:bg-white/10"}`}
            >
              <span
                className={`text-[8px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isZeroPogActivity ? "text-white/80" : "text-white/85"}`}
              >
                Opening Inv
              </span>
              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                {formatOverviewVal(row.lastQty, overviewUseMt).valueStr}
              </span>
            </div>
            <div
              className={`flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center transition-colors ${isZeroPogActivity ? "hover:bg-white/15" : "hover:bg-white/10"}`}
            >
              <span
                className={`text-[8px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isZeroPogActivity ? "text-white/80" : "text-white/85"}`}
              >
                End of Inv
              </span>
              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                {formatOverviewVal(row.currentQty, overviewUseMt).valueStr}
              </span>
            </div>
          </div>

          {/* Table 2: Stock in, idle stock, POG */}
          <div
            className={`flex-[3] flex divide-x rounded-[14px] overflow-hidden ${
              isZeroPogActivity
                ? "divide-red-200 bg-red-400 shadow-[0_12px_32px_rgba(239,68,68,0.18)]"
                : "divide-white/20 bg-primary shadow-[0_12px_32px_rgba(21,75,226,0.35)]"
            }`}
          >
            <div
              className={`flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center transition-colors ${isZeroPogActivity ? "hover:bg-white/15" : "hover:bg-white/10"}`}
            >
              <span
                className={`text-[8px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isZeroPogActivity ? "text-white/80" : "text-white/85"}`}
              >
                Stock in
              </span>
              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                {formatOverviewVal(row.sellIn, overviewUseMt).valueStr}
              </span>
            </div>
            <div
              className={`flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center transition-colors ${isZeroPogActivity ? "hover:bg-white/15" : "hover:bg-white/10"}`}
            >
              <span
                className={`text-[8px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isZeroPogActivity ? "text-white/80" : "text-amber-200"}`}
              >
                idle stock
              </span>
              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                {formatOverviewVal(row.idleStock, overviewUseMt).valueStr}
              </span>
            </div>
            <div
              className={`flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center transition-colors ${isZeroPogActivity ? "hover:bg-white/15" : "hover:bg-white/10"}`}
            >
              <span
                className={`text-[8px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isZeroPogActivity ? "text-white/80" : "text-cyan-250"}`}
              >
                POG
              </span>
              <span className="font-semibold text-[10.5px] truncate w-full text-white font-extrabold">
                {formatOverviewVal(row.pog, overviewUseMt).valueStr}
              </span>
            </div>
          </div>
        </div>

        {isExpanded && hasChildren && (
          <div className="pb-2 border-t border-[#edecff] bg-slate-50/15">
            {row.children.map((child: any) =>
              renderRecursivePogRow(child, depth + 1),
            )}
          </div>
        )}
      </div>
    );
  };

  if (showLoader) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen supports-[min-height:100dvh]:min-h-[100dvh] bg-[#f8fafc] px-4 pb-20 animate-in fade-in duration-500">
        <div className="bg-white p-8 rounded-[32px] shadow-[0_24px_64px_rgba(24,26,44,0.06)] border border-[#e2e8f0] flex flex-col items-center max-w-sm w-full gap-6">
          <div className="relative">
            <div className="size-24 border-4 border-[#edecff] rounded-full flex items-center justify-center">
              <span className="text-xl font-extrabold text-primary select-none">
                {loadProgress}%
              </span>
            </div>
            <div className="size-24 border-4 border-primary border-t-transparent rounded-full animate-spin absolute inset-0"></div>
          </div>

          <div className="text-center w-full">
            <h3 className="text-[#181a2c] font-black text-xl mb-1 select-none">
              Menyiapkan Data...
            </h3>
            <p className="text-[#8E94B7] text-[10px] font-bold uppercase tracking-widest select-none">
              SINKRONISASI DATABASE
            </p>
          </div>

          <div className="w-full flex flex-col items-center gap-2">
            <div className="w-full bg-slate-100 rounded-full h-2.5 overflow-hidden p-[1px] border border-[#f1f5f9]">
              <div
                className="bg-gradient-to-r from-primary to-cyan-400 h-full rounded-full transition-all duration-150 ease-out shadow-[0_2px_8px_rgba(21,75,226,0.25)]"
                style={{ width: `${loadProgress}%` }}
              ></div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  const renderPogKpiCard = (isSummaryTab: boolean) => {
    const dataEmpty = isSummaryTab
      ? filteredSummaryData.length === 0
      : aggregatedPogData.length === 0;
    if (dataEmpty) return null;
    
    const currentTotal = isSummaryTab ? totalSummaryPog : totalPog;
    const isFilterOpen = isSummaryTab ? isSummaryFilterOpen : isPogFilterOpen;
    const toggleFilter = () =>
      isSummaryTab
        ? setIsSummaryFilterOpen(!isSummaryFilterOpen)
        : setIsPogFilterOpen(!isPogFilterOpen);

    return (
      <div
        onClick={() => {
          if (window.innerWidth < 768) {
            toggleFilter();
          }
        }}
        className="flex-1 w-full min-w-0 bg-gradient-to-br from-primary to-cyan-400 p-5 md:px-7 rounded-[36px] shadow-[0_12px_32px_rgba(21,75,226,0.35)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.45)] hover:scale-[1.01] active:scale-[0.99] transition-all duration-250 cursor-pointer md:cursor-default select-none text-white mb-1.5 md:mb-0 md:hover:scale-100 md:active:scale-100"
      >
        <div className="flex justify-between items-center mb-4">
          <div className="flex items-center gap-1.5">
            <span className="material-symbols-outlined text-[18px] text-white/80">
              tune
            </span>
            <div className="flex flex-col">
              <span className="font-semibold text-sm uppercase tracking-wider text-white/90">
                Grand Total
              </span>
              {isSummaryTab && (
                <span className="text-[10px] font-bold text-white/70 uppercase tracking-widest mt-0.5">
                  {filterBelowCrop || "All"}
                </span>
              )}
            </div>
          </div>
          <div className="flex flex-col items-end">
            <span className="font-bold text-xl xl:text-2xl text-white">
              {formatOverviewVal(isSummaryTab ? totalSummary.selectedTotal : currentTotal.pog, overviewUseMt).valueStr}
            </span>
            <span className="text-[8px] text-white/80 uppercase tracking-widest font-bold">
              {isSummaryTab ? `Total ${formatOverviewVal(totalSummary.selectedTotal, overviewUseMt).unit}` : `POG (${formatOverviewVal(currentTotal.pog, overviewUseMt).unit})`}
            </span>
          </div>
        </div>

        {isSummaryTab ? (
          <div className="flex w-full divide-x divide-white/15 border-t border-white/15 pt-4 mt-2">
            {ALL_CLUSTER_KEYS.map((clusterKey) => {
              if (
                clusterKey === "Uncategorized" &&
                (!totalSummary[clusterKey] || totalSummary[clusterKey] === 0)
              )
                return null;
              const clusterConfig = CLUSTER_CONFIG.find(
                (c) => c.key === clusterKey,
              );
              const isSelected = selectedClusters.includes(clusterKey);
              return (
                <div
                  key={clusterKey}
                  onClick={(e) => {
                    e.stopPropagation();
                    setSelectedClusters((prev) =>
                      prev.includes(clusterKey)
                        ? prev.filter((k) => k !== clusterKey)
                        : [...prev, clusterKey]
                    );
                  }}
                  className={`flex-1 min-w-0 py-1.5 px-0.5 flex flex-col items-center justify-center text-center cursor-pointer transition-all ${isSelected ? "opacity-100 hover:bg-white/10 rounded-md" : "opacity-40 hover:opacity-60"}`}
                >
                  <span className="text-[9px] xl:text-[10px] font-bold uppercase tracking-wider text-white/85 mb-1 truncate w-full">
                    {clusterConfig?.label || clusterKey}
                  </span>
                  <span className="font-bold text-sm xl:text-base truncate w-full text-white">
                    {formatOverviewVal(totalSummary[clusterKey], overviewUseMt).valueStr}
                  </span>
                </div>
              );
            })}
          </div>
        ) : (
          <div className="flex flex-row w-full gap-1.5 md:gap-2 border-t border-white/15 pt-3.5 mt-1">
            {/* Table 1: Opening and End Inv group */}
            <div className="flex-[2] flex divide-x divide-white/15 bg-white/5 rounded-[12px] p-0.5">
              <div className="flex-1 min-w-0 py-1.5 px-0.5 flex flex-col items-center justify-center text-center">
                <span className="text-[8px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                  Opening
                </span>
                <span className="font-bold text-xs truncate w-full text-white">
                  {formatOverviewVal(currentTotal.lastQty, overviewUseMt).valueStr}
                </span>
              </div>
              <div className="flex-1 min-w-0 py-1.5 px-0.5 flex flex-col items-center justify-center text-center">
                <span className="text-[8px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                  End Inv
                </span>
                <span className="font-bold text-xs truncate w-full text-white">
                  {formatOverviewVal(currentTotal.currentQty, overviewUseMt).valueStr}
                </span>
              </div>
            </div>

            {/* Table 2: Stock in, idle stock, POG group */}
            <div className="flex-[3] flex divide-x divide-white/20 bg-gradient-to-br from-white/25 via-white/15 to-white/25 border border-white/20 rounded-[12px] p-0.5 shadow-[inset_0_1px_1px_rgba(255,255,255,0.25)]">
              <div className="flex-1 min-w-0 py-1.5 px-0.5 flex flex-col items-center justify-center text-center">
                <span className="text-[8px] font-bold uppercase tracking-wider text-white/95 mb-0.5 truncate w-full">
                  Stock In
                </span>
                <span className="font-extrabold text-xs truncate w-full text-white">
                  {formatOverviewVal(currentTotal.sellIn, overviewUseMt).valueStr}
                </span>
              </div>
              <div className="flex-1 min-w-0 py-1.5 px-0.5 flex flex-col items-center justify-center text-center hover:bg-white/5 transition-colors rounded-[12px]">
                <span className="text-[8px] font-bold uppercase tracking-wider text-amber-200 mb-0.5 truncate w-full">
                  Idle Stock
                </span>
                <span className="font-extrabold text-xs truncate w-full text-amber-100">
                  {formatOverviewVal(currentTotal.idleStock, overviewUseMt).valueStr}
                </span>
              </div>
              <div className="flex-1 min-w-0 py-1.5 px-0.5 flex flex-col items-center justify-center text-center hover:bg-white/5 transition-colors rounded-[12px]">
                <span className="text-[8px] font-bold uppercase tracking-wider text-cyan-200 mb-0.5 truncate w-full">
                  POG
                </span>
                <span className="font-black text-xs truncate w-full text-cyan-50">
                  {formatOverviewVal(currentTotal.pog, overviewUseMt).valueStr}
                </span>
              </div>
            </div>
          </div>
        )}

        <div className="border-t border-white/15 pt-5 mt-4 w-full">
          <div className="grid grid-cols-2 gap-3 sm:gap-6 lg:gap-8 items-start justify-center">
            {/* CHART 1: VISIT COVERAGE */}
            <div className="flex flex-col items-center text-center p-1 sm:p-3">
              <h4 className="text-[8.5px] sm:text-[10px] font-bold uppercase tracking-wider text-white/90 mb-3 sm:mb-4 font-sans line-clamp-1">
                Kunjungan (Visit)
              </h4>
              <div className="relative flex items-center justify-center shrink-0 mb-3 sm:mb-4 animate-in zoom-in duration-500">
                <svg
                  viewBox="0 0 220 220"
                  className="rotate-[-90deg] w-24 h-24 sm:w-32 sm:h-32 md:w-36 md:h-36 lg:w-40 lg:h-40 xl:w-44 xl:h-44"
                >
                  <circle
                    stroke="rgba(255, 255, 255, 0.15)"
                    fill="transparent"
                    strokeWidth={16}
                    r={95}
                    cx={110}
                    cy={110}
                  />
                  <circle
                    stroke="white"
                    fill="transparent"
                    strokeWidth={16}
                    strokeDasharray={`${2 * Math.PI * 95} ${2 * Math.PI * 95}`}
                    strokeDashoffset={
                      2 * Math.PI * 95 -
                      (activeVisitStats.percentage / 100) * (2 * Math.PI * 95)
                    }
                    style={{
                      strokeDashoffset: `${2 * Math.PI * 95 - (activeVisitStats.percentage / 100) * (2 * Math.PI * 95)}px`,
                    }}
                    r={95}
                    cx={110}
                    cy={110}
                    strokeLinecap="round"
                    className="transition-[stroke-dashoffset] duration-700 ease-out"
                  />
                </svg>
                <div className="absolute flex flex-col items-center justify-center">
                  <span className="text-lg sm:text-2xl md:text-3xl xl:text-4xl font-extrabold text-white tracking-tight leading-none">
                    {activeVisitStats.percentage}%
                  </span>
                  <span className="text-[7px] sm:text-[8px] md:text-[9px] font-bold uppercase tracking-wider text-white/70 mt-0.5 sm:mt-1">
                    Visited
                  </span>
                </div>
              </div>
              <p className="text-[9.5px] sm:text-[11px] md:text-[11.5px] leading-relaxed text-white/90 max-w-sm">
                {!mappingPic ||
                cleanForMatch(mappingPic) === "allteam" ||
                cleanForMatch(mappingPic) === "all_team"
                  ? "Team"
                  : normalizeName(mappingPic)}{" "}
                sudah melakukan visit sebanyak{" "}
                <span className="font-bold text-white">
                  {activeVisitStats.visited}
                </span>{" "}
                dari{" "}
                <span className="font-bold text-white">
                  {activeVisitStats.total}
                </span>{" "}
                Partner{" "}
                <span className="text-white/80">
                  ({activeVisitStats.percentage}%)
                </span>
              </p>
            </div>

            {/* CHART 2: CATEGORY DISTRIBUTION */}
            <div className="flex flex-col items-center text-center p-1 sm:p-3 border-l border-white/15">
              <h4 className="text-[8.5px] sm:text-[10px] font-bold uppercase tracking-wider text-white/90 mb-3 sm:mb-4 font-sans line-clamp-1">
                Kategori Partner
              </h4>
              <div className="relative flex items-center justify-center shrink-0 mb-3 sm:mb-4 md:mb-5 animate-in zoom-in duration-500">
                <svg
                  viewBox="0 0 220 220"
                  className="rotate-[-90deg] w-24 h-24 sm:w-32 sm:h-32 md:w-36 md:h-36 lg:w-40 lg:h-40 xl:w-44 xl:h-44"
                >
                  <circle
                    stroke="rgba(255, 255, 255, 0.1)"
                    fill="transparent"
                    strokeWidth={16}
                    r={95}
                    cx={110}
                    cy={110}
                  />
                  {categoryChartSegments.map((segment) => (
                    <circle
                      key={segment.name}
                      stroke={segment.color}
                      fill="transparent"
                      strokeWidth={16}
                      strokeLinecap="round"
                      strokeDasharray={segment.strokeDasharray}
                      strokeDashoffset={segment.strokeDashoffset}
                      style={{
                        strokeDashoffset: `${segment.strokeDashoffset}px`,
                      }}
                      r={95}
                      cx={110}
                      cy={110}
                      className="transition-[stroke-dashoffset] duration-700 ease-out"
                    />
                  ))}
                </svg>
                <div className="absolute flex flex-col items-center justify-center">
                  <span className="text-lg sm:text-2xl md:text-3xl xl:text-4xl font-extrabold text-white tracking-tight leading-none font-sans">
                    {categoryDistributionStats.total}
                  </span>
                  <span className="text-[7px] sm:text-[8px] md:text-[9px] font-bold uppercase tracking-wider text-white/70 mt-0.5 sm:mt-1 font-mono">
                    Partners
                  </span>
                </div>
              </div>
              <div className="w-full mt-1.5">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-1 px-0.5 sm:px-1 text-[8.5px] sm:text-[10px] text-white/90">
                  {categoryChartSegments.map((segment) => (
                    <div
                      key={segment.name}
                      className="flex items-center gap-1 sm:gap-1.5 bg-black/12 py-1 sm:py-1.5 px-1.5 sm:px-2 rounded-full border border-black/5 truncate hover:bg-black/25 transition-colors shadow-sm animate-in fade-in-50 duration-300"
                    >
                      <span
                        className="inline-block w-2 sm:w-2.5 h-2 sm:h-2.5 rounded-full shrink-0"
                        style={{ backgroundColor: segment.color }}
                      />
                      <span className="font-semibold text-white/90 truncate flex-1 text-left">
                        {segment.name}
                      </span>
                      <span className="font-mono font-bold text-white shrink-0">
                        {segment.value}{" "}
                        <span className="text-[7px] sm:text-[8px] font-normal text-white/75">
                          ({Math.round(segment.percentage)}%)
                        </span>
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderGrandTotalCard = () => {
    if (filteredSummaryData.length === 0) return null;
    return (
      <div
        className="flex-1 w-full min-w-0 bg-gradient-to-br from-primary to-cyan-400 p-4 md:py-4 md:px-6 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.35)] transition-all duration-250 select-none text-white flex flex-col gap-3 lg:gap-4"
      >
        {/* Left Side: Grand Total & Clusters */}
        <div className="flex flex-col w-full shrink-0 justify-center">
          <div className="flex justify-between items-center mb-2">
            <div className="flex items-center gap-2">
              <span className="material-symbols-outlined text-[18px] text-white/80">
                tune
              </span>
              <div className="flex flex-col mr-2">
                <span className="font-semibold text-xs uppercase tracking-wider text-white/90">
                  Grand Total
                </span>
                <span className="text-[9px] font-bold text-white/70 uppercase tracking-widest mt-0.5">
                  {filterBelowCrop || "All"}
                </span>
              </div>
              <div className="flex bg-white/10 p-0.5 rounded-lg border border-white/10 ml-2">
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setGrandTotalViewBy("hybrid");
                  }}
                  className={`px-3 py-1 rounded-md text-[9px] font-bold uppercase tracking-wider transition-colors ${
                    grandTotalViewBy === "hybrid"
                      ? "bg-white text-primary shadow-sm"
                      : "text-white/70 hover:text-white"
                  }`}
                >
                  Hybrid
                </button>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setGrandTotalViewBy("area");
                  }}
                  className={`px-3 py-1 rounded-md text-[9px] font-bold uppercase tracking-wider transition-colors ${
                    grandTotalViewBy === "area"
                      ? "bg-white text-primary shadow-sm"
                      : "text-white/70 hover:text-white"
                  }`}
                >
                  Province
                </button>
              </div>
            </div>
            <div className="flex flex-col items-end">
              <span className="font-bold text-xl xl:text-2xl text-white tracking-tight">
                {formatOverviewVal(totalSummary.selectedTotal, overviewUseMt).valueStr}
              </span>
              <span className="text-[8px] xl:text-[9px] text-white/80 uppercase tracking-widest font-bold">
                Total {formatOverviewVal(totalSummary.selectedTotal, overviewUseMt).unit}
              </span>
            </div>
          </div>
          <div className="flex w-full divide-x divide-white/15 border-t border-white/15 pt-2.5 mt-1">
            {ALL_CLUSTER_KEYS.map((clusterKey) => {
              if (
                clusterKey === "Uncategorized" &&
                (!totalSummary[clusterKey] || totalSummary[clusterKey] === 0)
              )
                return null;
              const clusterConfig = CLUSTER_CONFIG.find(
                (c) => c.key === clusterKey,
              );
              const isSelected = selectedClusters.includes(clusterKey);
              return (
                <div
                  key={clusterKey}
                  onClick={(e) => {
                    e.stopPropagation();
                    setSelectedClusters((prev) =>
                      prev.includes(clusterKey)
                        ? prev.filter((k) => k !== clusterKey)
                        : [...prev, clusterKey]
                    );
                  }}
                  className={`flex-1 min-w-0 py-1 px-0.5 flex flex-col items-center justify-center text-center cursor-pointer transition-all ${isSelected ? "opacity-100 hover:bg-white/10 rounded-md" : "opacity-40 hover:opacity-60"}`}
                >
                  <span className="text-[9px] xl:text-[10px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                    {clusterConfig?.label || clusterKey}
                  </span>
                  <span className="font-bold text-xs xl:text-sm truncate w-full text-white">
                    {formatOverviewWithUnit(totalSummary[clusterKey], overviewUseMt)}
                  </span>
                </div>
              );
            })}
          </div>
        </div>

        {/* Right Side: Total per Item & Aging */}
        <div className="flex flex-col flex-1 border-t border-white/15 pt-3">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-3 gap-2.5 w-full max-h-[300px] lg:max-h-none overflow-y-auto pr-1 custom-scrollbar">
            {grandTotalStats.map((stat, index) => (
              <div
                key={stat.name}
                className="bg-white/12 hover:bg-white/15 transition-all border border-white/25 py-2 px-2 rounded-[16px] flex flex-col shadow-sm"
              >
                <div className="flex items-center gap-2 mb-2">
                  <div className="size-6 rounded-lg bg-white/25 flex items-center justify-center text-white font-black text-xs shrink-0">
                    {index + 1}
                  </div>
                  <div className="min-w-0 flex-1 flex justify-between items-center gap-2">
                    <p className="text-xs md:text-[13px] font-bold text-white uppercase tracking-wider truncate leading-none" title={stat.name}>
                      {stat.name}
                    </p>
                    <p className="text-xs md:text-[13px] font-extrabold text-white tracking-tight leading-none shrink-0">
                      {formatOverviewWithUnit(stat.total, overviewUseMt)}
                    </p>
                  </div>
                </div>
                <div className="flex w-full divide-x divide-white/20 border-t border-white/20 pt-2 mt-1">
                  {selectedClusters.map((clusterKey) => {
                    const val = stat.clusters[clusterKey] || 0;
                    if (clusterKey === "Uncategorized" && val === 0)
                      return null;
                    const clusterConfig = CLUSTER_CONFIG.find(
                      (c) => c.key === clusterKey,
                    );
                    return (
                      <div
                        key={clusterKey}
                        className="flex-1 min-w-0 px-0.5 flex flex-col items-center justify-center text-center"
                      >
                        <span className="text-[9px] md:text-[10px] font-bold uppercase tracking-wider text-white/80 mb-0.5 truncate w-full">
                          {clusterConfig?.label || clusterKey}
                        </span>
                        <span className="font-bold text-[11px] md:text-[12px] truncate w-full text-white">
                          {formatOverviewVal(val, overviewUseMt).valueStr}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}
            {grandTotalStats.length === 0 && (
              <div className="col-span-full py-6 text-center text-white/60 text-xs font-medium">
                Tidak ada data.
              </div>
            )}
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="w-full max-w-2xl mx-auto lg:max-w-5xl xl:max-w-6xl px-5 pb-8 relative">
      <div className="flex items-stretch gap-2 mb-8 mt-6 -ml-2.5">
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            e.preventDefault();
            if (!isFetchingData) {
              fetchWorkingData();
            }
          }}
          disabled={isFetchingData}
          className="flex md:hidden flex-col items-center justify-center shrink-0 bg-gradient-to-r from-primary to-cyan-400 rounded-[24px] px-5 py-3 shadow-[0_12px_32px_rgba(21,75,226,0.35)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.45)] hover:scale-[1.02] active:scale-[0.98] transition-all min-w-[108px] cursor-pointer disabled:opacity-80 disabled:cursor-not-allowed select-none text-center border-0 appearance-none"
          title="Klik untuk Ambil Ulang Data dari Database"
        >
          <AdvantaLogo
            className={`size-14 text-white ${isFetchingData ? "animate-spin" : ""}`}
          />
          <span className="text-[9px] font-bold text-white/90 uppercase tracking-widest mt-1 leading-none">
            {isFetchingData ? "Refreshed" : "Radar"}
          </span>
        </button>
        <div className="flex md:hidden bg-gradient-to-r from-primary to-cyan-400 p-4 rounded-l-[28px] rounded-r-none -mr-5 shadow-[0_12px_32px_rgba(21,75,226,0.35)] text-white items-center justify-between flex-1 relative overflow-visible transition-all select-none duration-200">
          {/* Background and Logout wrapper that clips the huge user icon */}
          <div className="absolute inset-0 overflow-hidden rounded-l-[28px] z-0">
            <div
              onClick={(e) => {
                e.stopPropagation();
                e.preventDefault();
                setIsLogoutModalOpen(true);
              }}
              className="absolute inset-0 flex items-center justify-center opacity-10 cursor-pointer hover:opacity-15 active:opacity-20 transition-opacity"
              title="Klik untuk Keluar"
            >
              <UserIcon className="h-[260%] w-auto max-w-none object-contain object-center" />
            </div>
          </div>

          <div className="flex-1 relative z-10 min-w-0 pr-2 flex flex-col gap-2">
            <div className="flex flex-col gap-1.5">
              <div className="flex items-center gap-3">
                {greeting.imageUrl && (
                  <img
                    src={greeting.imageUrl}
                    referrerPolicy="no-referrer"
                    className="size-12 md:size-14 object-contain opacity-100 shrink-0 pointer-events-none select-none animate-rotate-sway"
                    alt=""
                  />
                )}
                <div className="flex-1 min-w-0 py-0.5">
                  <p className="text-[9px] text-white/80 font-semibold uppercase tracking-widest leading-none mb-0.5">
                    {greeting.text},
                  </p>
                  <h2 className="text-md md:text-lg font-bold text-white leading-tight truncate">
                    {userData.name}
                  </h2>
                </div>
              </div>
              <div className="flex items-center gap-1.5 flex-wrap">
                <span className="text-[10px] md:text-xs font-bold text-white uppercase bg-white/20 px-2.5 py-1 rounded-full border border-white/20 backdrop-blur-sm truncate select-none">
                  {userData.position}
                </span>
                {userData.province && userData.province !== "-" && (
                  <span className="text-[10px] md:text-xs font-bold text-white/80 uppercase bg-white/10 px-2.5 py-1 rounded-full border border-white/10 backdrop-blur-sm">
                    {userData.province}
                  </span>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>

      {activeTab === "overview" && (
        <div className="animate-in fade-in slide-in-from-right-4 duration-300">
          {/* Header */}
          <div className="mb-2 ml-1 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="size-10 rounded-2xl bg-[#154be2]/10 flex items-center justify-center shrink-0 border border-[#154be2]/20">
                <User className="size-5 text-[#154be2]" />
              </div>
              <div>
                <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
                  Executive{" "}
                  <span className="text-primary font-bold">Overview</span>
                </h1>
                <p className="text-[11px] text-[#8E94B7] font-semibold uppercase tracking-wider mt-0.5">
                  Analisis Kinerja & Pemantauan Tingkat Nasional
                </p>
              </div>
            </div>

            {/* Metrik Selector Buttons (Sejajar dengan Title, Rata Kanan) */}
            <div className="flex items-center gap-2">
              <div className="flex bg-slate-100 p-1 rounded-2xl border border-[#e2e8f0]">
                <button
                  onClick={() => setOverviewMetricFilter("activity")}
                  className={`px-6 py-3 rounded-xl text-sm font-black transition-all duration-200 ${
                    overviewMetricFilter === "activity"
                      ? "bg-[#154be2] text-white shadow-[0_4px_12px_rgba(21,75,226,0.25)]"
                      : "text-[#5c648e] hover:text-[#181a2c] hover:bg-slate-200"
                  }`}
                >
                  Activity
                </button>
                <button
                  onClick={() => setOverviewMetricFilter("nominal")}
                  className={`px-6 py-3 rounded-xl text-sm font-black transition-all duration-200 ${
                    overviewMetricFilter === "nominal"
                      ? "bg-[#154be2] text-white shadow-[0_4px_12px_rgba(21,75,226,0.25)]"
                      : "text-[#5c648e] hover:text-[#181a2c] hover:bg-slate-200"
                  }`}
                >
                  Nominal
                </button>
                <button
                  onClick={() => setOverviewMetricFilter("reach")}
                  className={`px-6 py-3 rounded-xl text-sm font-black transition-all duration-200 ${
                    overviewMetricFilter === "reach"
                      ? "bg-[#154be2] text-white shadow-[0_4px_12px_rgba(21,75,226,0.25)]"
                      : "text-[#5c648e] hover:text-[#181a2c] hover:bg-slate-200"
                  }`}
                >
                  Reach
                </button>
              </div>
            </div>
          </div>

          {/* Floating Filter Button */}
          <button
            onClick={() => setIsOverviewFilterOpen(!isOverviewFilterOpen)}
            className="fixed bottom-[110px] right-3 lg:bottom-10 lg:right-4 z-[60] bg-gradient-to-br from-[#154be2] to-[#123ebd] text-white p-3.5 lg:p-4 rounded-full shadow-[0_12px_32px_rgba(21,75,226,0.35)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.45)] hover:scale-105 active:scale-95 transition-all duration-300"
          >
            <span className="material-symbols-outlined text-[24px] lg:text-[28px]">filter_alt</span>
          </button>

          {/* Floating Detailed Filters */}
          <div
            className={`fixed inset-0 z-[70] transition-all duration-300 flex items-center justify-center p-4 ${isOverviewFilterOpen ? "opacity-100 visible pointer-events-auto" : "opacity-0 invisible pointer-events-none"}`}
          >
            {/* Backdrop */}
            <div 
              className="absolute inset-0 bg-black/40 backdrop-blur-sm transition-opacity duration-300" 
              onClick={() => setIsOverviewFilterOpen(false)} 
            />
            
            {/* Filter Content */}
            <div className={`relative w-full max-w-4xl bg-white p-6 md:p-8 rounded-[32px] shadow-[0_24px_64px_rgba(0,0,0,0.25)] transition-all duration-300 transform ${isOverviewFilterOpen ? "scale-100 translate-y-0" : "scale-95 translate-y-4"}`}>
              <div className="flex items-center justify-between gap-4 mb-6 pb-4 border-b border-slate-100">
                <div className="flex items-center gap-2">
                  <span className="material-symbols-outlined text-primary text-[20px] lg:text-[24px] font-semibold">
                    filter_alt
                  </span>
                  <h4 className="text-xs lg:text-sm font-bold text-[#181a2c] uppercase tracking-wider">
                    Filter Dashboard Analisis
                  </h4>
                </div>
                <button 
                  onClick={() => setIsOverviewFilterOpen(false)}
                  className="bg-slate-100 hover:bg-slate-200 text-slate-500 rounded-full w-8 h-8 flex items-center justify-center transition-colors"
                >
                  <span className="material-symbols-outlined text-[20px]">close</span>
                </button>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-7 gap-4 lg:gap-5">
                {/* Filter 2 - Month (Multi-Select Dropdown) */}
                <div className="flex flex-col gap-2" ref={monthDropdownRef}>
                  <label className="text-[10px] lg:text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Bulan
                  </label>
                  <div className="relative">
                    <button
                      type="button"
                      onClick={() => setIsMonthDropdownOpen(!isMonthDropdownOpen)}
                      className="w-full bg-[#fbfaff] border border-[#e2e8f0] rounded-xl px-2.5 lg:px-4 py-3 text-[11px] lg:text-xs font-bold text-[#181a2c] focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all text-left flex items-center justify-between shadow-sm cursor-pointer"
                    >
                      <span className="truncate pr-2">
                        {filterBelowMonth.length === 0
                          ? "Semua Bulan"
                          : filterBelowMonth.length === filterOptions.months.length
                          ? "Semua Bulan"
                          : filterBelowMonth.length <= 2
                          ? filterBelowMonth.join(", ")
                          : `${filterBelowMonth.length} Bulan Terpilih`}
                      </span>
                      <span 
                        className="material-symbols-outlined text-[18px] text-[#8E94B7] transition-transform duration-200"
                        style={{ transform: isMonthDropdownOpen ? 'rotate(180deg)' : 'none' }}
                      >
                        expand_more
                      </span>
                    </button>

                    {isMonthDropdownOpen && (
                      <div className="absolute left-0 mt-2 bg-white border border-[#e2e8f0] rounded-2xl shadow-xl z-50 max-h-60 overflow-y-auto p-2.5 flex flex-col gap-1.5 scrollbar-thin min-w-[185px] w-max max-w-[240px]">
                        <button
                          type="button"
                          onClick={() => {
                            if (filterBelowMonth.length > 0) {
                              setFilterBelowMonth([]);
                            } else {
                              setFilterBelowMonth([...filterOptions.months]);
                            }
                          }}
                          className="flex items-center gap-2.5 px-3 py-2 text-left hover:bg-slate-50 rounded-lg transition-colors cursor-pointer w-full text-xs font-semibold"
                        >
                          <input
                            type="checkbox"
                            checked={filterBelowMonth.length === filterOptions.months.length}
                            onChange={() => {}} // Handled by button click
                            className="rounded text-primary focus:ring-primary h-3.5 w-3.5 cursor-pointer accent-[#154be2]"
                          />
                          <span className="text-[#181a2c] font-extrabold uppercase tracking-wide text-[10px]">
                            {filterBelowMonth.length > 0 ? "Reset Pilihan" : "Pilih Semua"}
                          </span>
                        </button>
                        <div className="border-t border-slate-100 my-1"></div>
                        {filterOptions.months.map((m) => {
                          const isChecked = filterBelowMonth.includes(m);
                          return (
                            <button
                              key={m}
                              type="button"
                              onClick={() => {
                                if (isChecked) {
                                  setFilterBelowMonth(filterBelowMonth.filter((item) => item !== m));
                                } else {
                                  setFilterBelowMonth([...filterBelowMonth, m]);
                                }
                              }}
                              className="flex items-center gap-2.5 px-3 py-2 text-left hover:bg-slate-50 rounded-lg transition-colors cursor-pointer w-full text-xs"
                            >
                              <input
                                type="checkbox"
                                checked={isChecked}
                                onChange={() => {}} // Handled by button click
                                className="rounded text-primary focus:ring-primary h-3.5 w-3.5 cursor-pointer accent-[#154be2]"
                              />
                              <span className="text-slate-700 font-semibold">{m}</span>
                            </button>
                          );
                        })}
                      </div>
                    )}
                  </div>
                </div>

                {/* Filter 2 - Activity */}
                <div className="flex flex-col gap-2">
                  <label className="text-[10px] lg:text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Activity
                  </label>
                  <div className="relative">
                    <select
                      value={activeActivityFilter || "All"}
                      onChange={(e) => {
                        const val = e.target.value;
                        setActiveActivityFilter(val === "All" ? null : val);
                      }}
                      className="w-full bg-[#fbfaff] border border-[#e2e8f0] rounded-xl px-2.5 lg:px-4 py-3 text-[11px] lg:text-xs font-bold text-[#181a2c] focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all appearance-none cursor-pointer pr-10"
                    >
                      <option value="All">Semua Activity</option>
                      {(() => {
                        const REGULAR_LIST = ["FFD", "FM", "ODP", "SFT", "BC", "PT"];
                        const ADHOC_LIST = ["BFFD", "BFM", "BFT", "CRV", "EXP"];
                        let list = ["FFD", "FM", "ODP", "SFT", "BFFD", "BFM", "BFT", "BC", "CRV", "EXP", "PT"];
                        if (filterBelowType === "Regular") list = REGULAR_LIST;
                        if (filterBelowType === "AdHoc") list = ADHOC_LIST;
                        return list.map((act) => (
                          <option key={act} value={act}>
                            {act} - {getActivityFullName(act)}
                          </option>
                        ));
                      })()}
                    </select>
                    <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-[18px] text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>

                {/* Filter 2 - Material */}
                <div className="flex flex-col gap-2">
                  <label className="text-[10px] lg:text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Hybrid
                  </label>
                  <div className="relative">
                    <select
                      value={filterBelowMaterial}
                      onChange={(e) => setFilterBelowMaterial(e.target.value)}
                      className="w-full bg-[#fbfaff] border border-[#e2e8f0] rounded-xl px-2.5 lg:px-4 py-3 text-[11px] lg:text-xs font-bold text-[#181a2c] focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all appearance-none cursor-pointer pr-10"
                    >
                      <option value="All">Semua Hybrid</option>
                      {filterOptions.materials.map((m) => (
                        <option key={m} value={m}>
                          {m}
                        </option>
                      ))}
                    </select>
                    <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-[18px] text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>

                {/* Filter 2 - Team */}
                <div className="flex flex-col gap-2">
                  <label className="text-[10px] lg:text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Tim (PIC)
                  </label>
                  <div className="relative">
                    <select
                      value={filterBelowTeam}
                      onChange={(e) => setFilterBelowTeam(e.target.value)}
                      className="w-full bg-[#fbfaff] border border-[#e2e8f0] rounded-xl px-2.5 lg:px-4 py-3 text-[11px] lg:text-xs font-bold text-[#181a2c] focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all appearance-none cursor-pointer pr-10"
                    >
                      <option value="All">Semua PIC</option>
                      {filterOptions.teams.map((t) => (
                        <option key={t} value={t}>
                          {t}
                        </option>
                      ))}
                    </select>
                    <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-[18px] text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>

                {/* Filter 2 - Area */}
                <div className="flex flex-col gap-2">
                  <label className="text-[10px] lg:text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Wilayah
                  </label>
                  <div className="relative">
                    <select
                      value={filterBelowArea}
                      onChange={(e) => setFilterBelowArea(e.target.value)}
                      className="w-full bg-[#fbfaff] border border-[#e2e8f0] rounded-xl px-2.5 lg:px-4 py-3 text-[11px] lg:text-xs font-bold text-[#181a2c] focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all appearance-none cursor-pointer pr-10"
                    >
                      <option value="All">Semua Wilayah</option>
                      {filterOptions.areas.map((a) => (
                        <option key={a} value={a}>
                          {a}
                        </option>
                      ))}
                    </select>
                    <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-[18px] text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>

                {/* Filter 2 - Crop */}
                <div className="flex flex-col gap-2">
                  <label className="text-[10px] lg:text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Komoditas
                  </label>
                  <div className="relative">
                    <select
                      value={filterBelowCrop}
                      onChange={(e) => setFilterBelowCrop(e.target.value)}
                      className="w-full bg-[#fbfaff] border border-[#e2e8f0] rounded-xl px-2.5 lg:px-4 py-3 text-[11px] lg:text-xs font-bold text-[#181a2c] focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all appearance-none cursor-pointer pr-10"
                    >
                      <option value="All">Semua Crop</option>
                      {["Field Corn", "Fresh Corn", "Vegetables"].map((crop) => (
                        <option key={crop} value={crop}>
                          {crop}
                        </option>
                      ))}
                    </select>
                    <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-[18px] text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>

                {/* Filter 2 - Type */}
                <div className="flex flex-col gap-2">
                  <label className="text-[10px] lg:text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Type
                  </label>
                  <div className="relative">
                    <select
                      value={filterBelowType}
                      onChange={(e) => setFilterBelowType(e.target.value)}
                      className="w-full bg-[#fbfaff] border border-[#e2e8f0] rounded-xl px-2.5 lg:px-4 py-3 text-[11px] lg:text-xs font-bold text-[#181a2c] focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all appearance-none cursor-pointer pr-10"
                    >
                      <option value="All">Regular & AdHoc</option>
                      <option value="Regular">Regular Only</option>
                      <option value="AdHoc">AdHoc Only</option>
                    </select>
                    <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-[18px] text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div className="mb-3.5">
            {renderGrandTotalCard()}
          </div>

          {/* Charts Grid Row 1 */}
          <div className="grid grid-cols-1 lg:grid-cols-11 gap-6 mb-6">
            {/* Chart 1: Sales (POG) & Stock per Area / Dimension */}
            <div className="lg:col-span-7 bg-white p-4 lg:p-5 rounded-[48px] shadow-[0_12px_32px_rgba(21,75,226,0.18)] border border-[#154be2]/8 flex flex-col gap-0 h-fit">
              <div className="flex flex-col gap-1 mb-1 pb-1 border-b border-[#f0effc]/60">
                <div className="flex items-start justify-between">
                  <div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => {
                          setFocusedChartType("main");
                          setIsChartFocusedModalOpen(true);
                        }}
                        className="p-1 hover:bg-[#154be2]/10 active:bg-[#154be2]/20 transition-all rounded-lg border border-[#154be2]/10 text-primary flex items-center justify-center shrink-0 cursor-pointer"
                        title="Fokus Grafik Utama"
                      >
                        <span className="material-symbols-outlined text-sm font-semibold">
                          {overviewGroupDimension === "material"
                            ? "widgets"
                            : overviewGroupDimension === "province"
                              ? "map"
                              : "analytics"}
                        </span>
                      </button>
                      <h3 className="text-xs font-bold text-[#181a2c] tracking-tight">
                        {overviewGroupDimension === "area"
                          ? "Budget Effectiveness Wilayah (Area)"
                          : overviewGroupDimension === "province"
                            ? "Budget Effectiveness per Provinsi"
                            : overviewGroupDimension === "sales_agronomist"
                              ? "Budget Effectiveness Sales Agronomist (SA)"
                              : overviewGroupDimension === "hybrid" || overviewGroupDimension === "material"
                                ? "Budget Effectiveness per Hybrid"
                                : "Budget Effectiveness per Activity"}
                      </h3>
                    </div>
                  </div>
                </div>

                <div className="flex flex-wrap items-center justify-between gap-3 w-full">
                  {/* Selector Filter 2: Dimensi Grouping */}
                  <div className="flex flex-wrap items-center gap-2">
                    <div className="flex items-center gap-2 bg-[#fbfaff] px-2.5 py-1 rounded-xl border border-[#e2e8f0]/80">
                      <span className="text-[9.5px] font-bold text-[#8E94B7] uppercase tracking-wider">
                        Dimensi:
                      </span>
                      <div className="relative">
                        <select
                          value={overviewGroupDimension}
                          onChange={(e: any) =>
                            setOverviewGroupDimension(e.target.value as any)
                          }
                          className="bg-transparent text-[10.5px] font-black text-[#154be2] focus:outline-none focus:ring-0 appearance-none cursor-pointer pr-6 py-0.5"
                        >
                          <option value="area">Area</option>
                          <option value="province">Province</option>
                          <option value="sales_agronomist">Sales Agronomist</option>
                          <option value="hybrid">Hybrids</option>
                          <option value="activity">Activity</option>
                        </select>
                        <span className="material-symbols-outlined absolute right-0 top-1/2 -translate-y-1/2 text-[14px] text-primary pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>

                    {/* Selector Filter: Sub Grouping */}
                    <div className="flex items-center gap-2 bg-[#fbfaff] px-2.5 py-1 rounded-xl border border-[#e2e8f0]/80">
                      <span className="text-[9.5px] font-bold text-[#8E94B7] uppercase tracking-wider">
                        Sub:
                      </span>
                      <div className="relative">
                        <select
                          value={subGroupDimension}
                          onChange={(e: any) =>
                            setSubGroupDimension(e.target.value as any)
                          }
                          className="bg-transparent text-[10.5px] font-black text-[#154be2] focus:outline-none focus:ring-0 appearance-none cursor-pointer pr-6 py-0.5"
                        >
                          <option value="area">Area</option>
                          <option value="province">Province</option>
                          <option value="sales_agronomist">Sales Agronomist</option>
                          <option value="hybrid">Hybrids</option>
                          <option value="activity">Activity</option>
                        </select>
                        <span className="material-symbols-outlined absolute right-0 top-1/2 -translate-y-1/2 text-[14px] text-primary pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>
                  </div>



                  {/* Legend aligned side-by-side */}
                  <div className="flex items-center gap-4 bg-[#fbfaff] px-3.5 py-1.5 rounded-xl border border-[#e2e8f0]/40 shrink-0 ml-auto select-none">
                    <button
                      type="button"
                      onClick={() => setShowBudgetBar(prev => !prev)}
                      className={`flex items-center gap-2 hover:opacity-85 transition-all cursor-pointer ${!showBudgetBar ? "opacity-35 line-through" : ""}`}
                      title="Klik untuk menyembunyikan/menampilkan Budget"
                    >
                      <span className="size-3.5 rounded-[4px] bg-gradient-to-tr from-[#154be2] to-[#3b82f6]" />
                      <span className="text-[12px] font-extrabold text-[#4e5572]">
                        Budget
                      </span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setShowActualBar(prev => !prev)}
                      className={`flex items-center gap-2 hover:opacity-85 transition-all cursor-pointer ${!showActualBar ? "opacity-35 line-through" : ""}`}
                      title="Klik untuk menyembunyikan/menampilkan Actual"
                    >
                      <span className="size-3.5 rounded-[4px] bg-gradient-to-tr from-[#06b6d4] to-[#22d3ee]" />
                      <span className="text-[12px] font-extrabold text-[#4e5572]">
                        Actual
                      </span>
                    </button>
                  </div>
                </div>
              </div>

              <div className="w-full overflow-x-auto scrollbar-thin select-none">
                <div
                  style={{
                    minWidth: `${Math.max(600, overviewStats.areaChartData.length * 95)}px`,
                    width: "100%",
                    height: "230px",
                  }}
                  className="font-sans"
                >
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={overviewStats.areaChartData}
                      margin={{ top: 38, right: 10, left: -10, bottom: 0 }}
                      barGap={currentBarGap}
                      barCategoryGap={currentBarCategoryGap}
                      onMouseMove={(state) => {
                        if (state && state.activeLabel) {
                          setHoveredLabel(state.activeLabel);
                          if (state.activeLabel !== dismissedTooltipLabel) {
                            setDismissedTooltipLabel(null);
                          }
                        } else {
                          setHoveredLabel(null);
                        }
                      }}
                      onMouseLeave={() => {
                        setHoveredLabel(null);
                      }}
                    >
                      <defs>
                        <linearGradient
                          id="colorAreaPog"
                          x1="0"
                          y1="0"
                          x2="0"
                          y2="1"
                        >
                          <stop
                            offset="0%"
                            stopColor="#154be2"
                            stopOpacity={0.95}
                          />
                          <stop
                            offset="100%"
                            stopColor="#3b82f6"
                            stopOpacity={0.7}
                          />
                        </linearGradient>
                        <linearGradient
                          id="colorAreaStock"
                          x1="0"
                          y1="0"
                          x2="0"
                          y2="1"
                        >
                          <stop
                            offset="0%"
                            stopColor="#06b6d4"
                            stopOpacity={1.0}
                          />
                          <stop
                            offset="100%"
                            stopColor="#22d3ee"
                            stopOpacity={1.0}
                          />
                        </linearGradient>
                      </defs>
                      <CartesianGrid
                        strokeDasharray="4 4"
                        vertical={false}
                        stroke="#e2e8f0"
                      />
                      <XAxis
                        dataKey="name"
                        tick={<CustomXAxisTick chartData={overviewStats.areaChartData} metricType={overviewMetricFilter} />}
                        axisLine={false}
                        tickLine={false}
                        interval={0}
                        height={65}
                      />
                      <YAxis
                        hide={true}
                        domain={[0, (dataMax: any) => (dataMax === 0 ? 100 : Math.round(dataMax * 1.25))]}
                        tick={{ fill: "#8E94B7", fontSize: 9, fontWeight: 500 }}
                        axisLine={false}
                        tickLine={false}
                      />
                      <Tooltip
                        cursor={{ fill: "rgba(21, 75, 226, 0.03)" }}
                        content={<CustomChartTooltip metricType={overviewMetricFilter} dismissedLabel={dismissedTooltipLabel} />}
                      />
                      <Bar
                        hide={!showBudgetBar}
                        dataKey={
                          overviewMetricFilter === "activity"
                            ? "budgetActivity"
                            : overviewMetricFilter === "reach"
                            ? "budgetReach"
                            : "budgetNominal"
                        }
                        name="Budget"
                        fill="url(#colorAreaPog)"
                        radius={[6, 6, 0, 0]}
                        maxBarSize={currentMaxBarSize}
                        background={<CustomBarBackground data={overviewStats.areaChartData} activeKey={activeMainBarKey} />}
                      >
                        {overviewStats.areaChartData.map((entry: any, index: number) => {
                          const isActive = activeMainBarKey === entry.name;
                          return (
                            <Cell
                              key={`cell-budget-${index}`}
                              cursor="pointer"
                              fill={isActive ? "#ea580c" : "url(#colorAreaPog)"}
                              fillOpacity={1.0}
                              onClick={() => {
                                clickedBarRef.current = true;
                                if (activeMainBarKey === entry.name) {
                                  setDismissedTooltipLabel(entry.name);
                                } else {
                                  setDismissedTooltipLabel(null);
                                }
                                setActiveMainBarKey(prev => prev === entry.name ? null : entry.name);
                              }}
                            />
                          );
                        })}
                        <LabelList
                          dataKey={
                            overviewMetricFilter === "activity"
                              ? "budgetActivity"
                              : overviewMetricFilter === "reach"
                              ? "budgetReach"
                              : "budgetNominal"
                          }
                          content={<CustomBudgetLabel metricType={overviewMetricFilter} />}
                        />
                      </Bar>
                      <Bar
                        hide={!showActualBar}
                        dataKey={
                          overviewMetricFilter === "activity"
                            ? "actualActivity"
                            : overviewMetricFilter === "reach"
                            ? "actualReach"
                            : "actualNominal"
                        }
                        name="Actual"
                        fill="url(#colorAreaStock)"
                        radius={[6, 6, 0, 0]}
                        maxBarSize={currentMaxBarSize}
                        background={<CustomBarBackground data={overviewStats.areaChartData} activeKey={activeMainBarKey} />}
                      >
                        {overviewStats.areaChartData.map((entry: any, index: number) => {
                          const isActive = activeMainBarKey === entry.name;
                          return (
                            <Cell
                              key={`cell-actual-${index}`}
                              cursor="pointer"
                              fill={isActive ? "#f97316" : "url(#colorAreaStock)"}
                              fillOpacity={1.0}
                              onClick={() => {
                                clickedBarRef.current = true;
                                if (activeMainBarKey === entry.name) {
                                  setDismissedTooltipLabel(entry.name);
                                } else {
                                  setDismissedTooltipLabel(null);
                                }
                                setActiveMainBarKey(prev => prev === entry.name ? null : entry.name);
                              }}
                            />
                          );
                        })}
                        <LabelList
                          dataKey={
                            overviewMetricFilter === "activity"
                              ? "actualActivity"
                              : overviewMetricFilter === "reach"
                              ? "actualReach"
                              : "actualNominal"
                          }
                          content={<CustomActualLabel metricType={overviewMetricFilter} />}
                        />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* Divider and Subheading for Sub Bar Chart */}
              <div className="mt-0.5 pt-0.5 border-t border-[#f0effc]/40 flex flex-col gap-0.5 shrink-0">
                <div className="flex items-center justify-between">
                  <div>
                    <div className="flex items-center gap-1.5 flex-wrap">
                      <button
                        onClick={() => {
                          setFocusedChartType("sub");
                          setIsChartFocusedModalOpen(true);
                        }}
                        className="p-1 hover:bg-[#06b6d4]/10 active:bg-[#06b6d4]/20 transition-all rounded-lg border border-[#06b6d4]/10 text-[#06b6d4] flex items-center justify-center shrink-0 cursor-pointer"
                        title="Fokus Grafik Sub"
                      >
                        <span className="material-symbols-outlined text-[13px] font-semibold">
                          bar_chart
                        </span>
                      </button>
                      <h3 className="text-[11px] font-bold text-[#181a2c] tracking-tight">
                        {subGroupDimension === "area"
                          ? "Sub Budget Effectiveness Wilayah (Area)"
                          : subGroupDimension === "province"
                            ? "Sub Budget Effectiveness per Provinsi"
                            : subGroupDimension === "sales_agronomist"
                              ? "Sub Budget Effectiveness Sales Agronomist (SA)"
                              : subGroupDimension === "hybrid" || subGroupDimension === "material"
                                ? "Sub Budget Effectiveness per Hybrid"
                                : "Sub Budget Effectiveness per Activity"}
                      </h3>
                      {activeMainBarKey && (
                        <button
                          onClick={() => setActiveMainBarKey(null)}
                          className="ml-auto text-[9px] font-bold text-[#154be2] bg-[#154be2]/10 hover:bg-[#154be2]/20 px-2 py-0.5 rounded-md transition-all flex items-center gap-1 cursor-pointer"
                          title="Click to clear filter"
                        >
                          <span>Filter: {activeMainBarKey}</span>
                          <span className="material-symbols-outlined text-[10px] font-bold">close</span>
                        </button>
                      )}
                      {activeActivityFilter && (
                        <button
                          onClick={() => setActiveActivityFilter(null)}
                          className={`${activeMainBarKey ? "ml-2" : "ml-auto"} text-[9px] font-bold text-amber-600 bg-amber-500/10 hover:bg-amber-500/20 px-2 py-0.5 rounded-md transition-all flex items-center gap-1 cursor-pointer`}
                          title="Click to clear Activity filter"
                        >
                          <span>Activity: {activeActivityFilter}</span>
                          <span className="material-symbols-outlined text-[10px] font-bold">close</span>
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              </div>

              {/* Second Bar Chart: Sub Group Dimension */}
              <div className="w-full overflow-x-auto scrollbar-thin select-none mt-0">
                <div
                  style={{
                    minWidth: `${Math.max(600, (overviewStats.subChartData?.length || 0) * 95)}px`,
                    width: "100%",
                    height: "230px",
                  }}
                  className="font-sans"
                >
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={overviewStats.subChartData || []}
                      margin={{ top: 38, right: 10, left: -10, bottom: 0 }}
                      barGap={currentBarGap}
                      barCategoryGap={currentBarCategoryGap}
                      onMouseMove={(state) => {
                        if (state && state.activeLabel) {
                          setHoveredLabel(state.activeLabel);
                          if (state.activeLabel !== dismissedSubTooltipLabel) {
                            setDismissedSubTooltipLabel(null);
                          }
                        } else {
                          setHoveredLabel(null);
                        }
                      }}
                      onMouseLeave={() => {
                        setHoveredLabel(null);
                      }}
                    >
                      <defs>
                        <linearGradient
                          id="colorSubPog"
                          x1="0"
                          y1="0"
                          x2="0"
                          y2="1"
                        >
                          <stop
                            offset="0%"
                            stopColor="#154be2"
                            stopOpacity={0.95}
                          />
                          <stop
                            offset="100%"
                            stopColor="#3b82f6"
                            stopOpacity={0.7}
                          />
                        </linearGradient>
                        <linearGradient
                          id="colorSubStock"
                          x1="0"
                          y1="0"
                          x2="0"
                          y2="1"
                        >
                          <stop
                            offset="0%"
                            stopColor="#06b6d4"
                            stopOpacity={1.0}
                          />
                          <stop
                            offset="100%"
                            stopColor="#22d3ee"
                            stopOpacity={1.0}
                          />
                        </linearGradient>
                      </defs>
                      <CartesianGrid
                        strokeDasharray="4 4"
                        vertical={false}
                        stroke="#e2e8f0"
                      />
                      <XAxis
                        dataKey="name"
                        tick={<CustomXAxisTick chartData={overviewStats.subChartData} metricType={overviewMetricFilter} />}
                        axisLine={false}
                        tickLine={false}
                        interval={0}
                        height={65}
                      />
                      <YAxis
                        hide={true}
                        domain={[0, (dataMax: any) => (dataMax === 0 ? 100 : Math.round(dataMax * 1.25))]}
                        tick={{ fill: "#8E94B7", fontSize: 9, fontWeight: 500 }}
                        axisLine={false}
                        tickLine={false}
                      />
                      <Tooltip
                        cursor={{ fill: "rgba(21, 75, 226, 0.03)" }}
                        content={<CustomChartTooltip metricType={overviewMetricFilter} dismissedLabel={dismissedSubTooltipLabel} />}
                      />
                      <Bar
                        hide={!showBudgetBar}
                        dataKey={
                          overviewMetricFilter === "activity"
                            ? "budgetActivity"
                            : overviewMetricFilter === "reach"
                            ? "budgetReach"
                            : "budgetNominal"
                        }
                        name="Budget"
                        fill="url(#colorSubPog)"
                        radius={[6, 6, 0, 0]}
                        maxBarSize={currentMaxBarSize}
                        background={<CustomBarBackground data={overviewStats.subChartData || []} activeKey={activeSubBarKey} />}
                      >
                        {(overviewStats.subChartData || []).map((entry: any, index: number) => {
                          const isActive = activeSubBarKey === entry.name;
                          return (
                            <Cell
                              key={`cell-sub-budget-${index}`}
                              cursor="pointer"
                              fill={isActive ? "#ea580c" : "url(#colorSubPog)"}
                              fillOpacity={1.0}
                              onClick={() => {
                                clickedBarRef.current = true;
                                if (activeSubBarKey === entry.name) {
                                  setDismissedSubTooltipLabel(entry.name);
                                } else {
                                  setDismissedSubTooltipLabel(null);
                                }
                                setActiveSubBarKey(prev => prev === entry.name ? null : entry.name);
                              }}
                            />
                          );
                        })}
                        <LabelList
                          dataKey={
                            overviewMetricFilter === "activity"
                              ? "budgetActivity"
                              : overviewMetricFilter === "reach"
                              ? "budgetReach"
                              : "budgetNominal"
                          }
                          content={<CustomBudgetLabel metricType={overviewMetricFilter} />}
                        />
                      </Bar>
                      <Bar
                        hide={!showActualBar}
                        dataKey={
                          overviewMetricFilter === "activity"
                            ? "actualActivity"
                            : overviewMetricFilter === "reach"
                            ? "actualReach"
                            : "actualNominal"
                        }
                        name="Actual"
                        fill="url(#colorSubStock)"
                        radius={[6, 6, 0, 0]}
                        maxBarSize={currentMaxBarSize}
                        background={<CustomBarBackground data={overviewStats.subChartData || []} activeKey={activeSubBarKey} />}
                      >
                        {(overviewStats.subChartData || []).map((entry: any, index: number) => {
                          const isActive = activeSubBarKey === entry.name;
                          return (
                            <Cell
                              key={`cell-sub-actual-${index}`}
                              cursor="pointer"
                              fill={isActive ? "#f97316" : "url(#colorSubStock)"}
                              fillOpacity={1.0}
                              onClick={() => {
                                clickedBarRef.current = true;
                                if (activeSubBarKey === entry.name) {
                                  setDismissedSubTooltipLabel(entry.name);
                                } else {
                                  setDismissedSubTooltipLabel(null);
                                }
                                setActiveSubBarKey(prev => prev === entry.name ? null : entry.name);
                              }}
                            />
                          );
                        })}
                        <LabelList
                          dataKey={
                            overviewMetricFilter === "activity"
                              ? "actualActivity"
                              : overviewMetricFilter === "reach"
                              ? "actualReach"
                              : "actualNominal"
                          }
                          content={<CustomActualLabel metricType={overviewMetricFilter} />}
                        />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            {/* KPI Cards Stack (Replacing Segmentasi Partner) */}
            <div className="lg:col-span-4 flex flex-col h-full min-h-0">
              <div className="flex flex-col gap-3.5 overflow-y-auto max-h-[460px] lg:max-h-[660px] lg:h-[660px] pl-4 pr-5 lg:pl-5 lg:pr-6.5 py-4.5 scrollbar-thin">
                {activityDonutCardsData.map((act) => {
                  const isSelected = (act.name === "TOTAL" && activeActivityFilter === null) || (activeActivityFilter === act.name);
                  const isNominal = overviewMetricFilter === "nominal";
                  const gap = act.actual - act.budget;
                  const gapSign = gap > 0 ? "+" : "";

                  let gapStr = "";
                  if (isNominal) {
                    const absGap = Math.abs(gap);
                    if (absGap >= 1000000000) gapStr = `${gapSign}${(gap / 1000000000).toFixed(1)}M`;
                    else if (absGap >= 1000000) gapStr = `${gapSign}${(gap / 1000000).toFixed(0)}Jt`;
                    else gapStr = `${gapSign}${gap.toLocaleString()}`;
                  } else {
                    gapStr = `${gapSign}${gap}`;
                  }
                  if (gap === 0) gapStr = "0";

                  return (
                    <button
                      key={act.name}
                      onClick={() => {
                        if (act.name === "TOTAL") {
                          setActiveActivityFilter(null);
                        } else {
                          setActiveActivityFilter((prev) => (prev === act.name ? null : act.name));
                        }
                      }}
                      className={`rounded-[32px] lg:w-[94%] lg:mx-auto flex flex-row items-center justify-between relative overflow-hidden group transition-all duration-300 w-full text-left cursor-pointer border-0 ${
                        act.name === "TOTAL"
                          ? "min-h-[116px] py-5 px-5 lg:py-6 lg:px-6"
                          : "min-h-[96px] py-4 px-4 lg:py-5 lg:px-4.5"
                      } ${
                        isSelected
                          ? "bg-gradient-to-r from-[#154be2] to-cyan-500 text-white shadow-[0_12px_32px_rgba(21,75,226,0.25)] scale-[1.02]"
                          : "bg-white shadow-[0_12px_32px_rgba(21,75,226,0.18)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.28)]"
                      }`}
                    >
                      <div className="flex flex-col justify-center z-10 min-w-0 flex-1 pr-2">
                        <div>
                          {/* Abbreviation above full name */}
                          <div className="mb-0.5">
                            <span className={`${act.name === "TOTAL" ? "text-[10px] px-2 py-0.5" : "text-[8px] px-1.5 py-0.2"} font-black rounded-md uppercase tracking-wider ${
                              isSelected
                                ? "bg-white/20 text-white border border-white/20"
                                : "bg-[#154be2]/10 text-[#154be2]"
                            }`}>
                              {act.name}
                            </span>
                          </div>
                          <h4 className={`${act.name === "TOTAL" ? "text-[16px] lg:text-[17px]" : "text-[12px] lg:text-[12.5px]"} font-black leading-tight truncate max-w-[180px] lg:max-w-[220px] ${
                            isSelected ? "text-white" : "text-[#181a2c]"
                          }`} title={act.fullName}>
                            {act.fullName}
                          </h4>
                        </div>
                        <div className="mt-1.5">
                          <div className={`font-sans font-extrabold ${act.name === "TOTAL" ? "text-[12.5px] lg:text-[13px]" : "text-[12.5px] lg:text-[13.5px]"} ${isSelected ? "text-white" : ""} flex items-center gap-1.5 flex-wrap`}>
                            <span className={isSelected ? "text-white font-black" : "text-[#06b6d4]"}>
                              {act.actualStr}
                            </span>
                            <span className={isSelected ? "text-white/50" : "text-[#8E94B7]"}>
                              /
                            </span>
                            <span className={isSelected ? "text-white/85" : "text-[#154be2]"}>
                              {act.budgetStr}
                            </span>
                            <span className={`ml-1.5 px-2 py-0.5 rounded font-black ${
                              act.name === "TOTAL"
                                ? "text-[12px] lg:text-[12.5px]"
                                : "text-[12.5px] lg:text-[13.5px]"
                            } ${
                              isSelected
                                ? "bg-white/20 text-white"
                                : gap >= 0
                                  ? "bg-emerald-50 text-emerald-700 border border-emerald-100/50"
                                  : "bg-rose-50 text-rose-700 border border-rose-100/50"
                            }`}>
                              {gapStr}
                            </span>
                          </div>
                        </div>
                      </div>
   
                      {/* Enriched Donut Chart Wrapper */}
                      <div className="flex flex-col items-center justify-center shrink-0 z-10 pl-1">
                        {/* Donut Chart Visual */}
                        <div className={`relative shrink-0 ${act.name === "TOTAL" ? "size-22" : "size-18"}`}>
                          <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                              <Pie
                                data={act.chartData}
                                cx="50%"
                                cy="50%"
                                innerRadius={act.name === "TOTAL" ? 28 : 22}
                                outerRadius={act.name === "TOTAL" ? 38 : 31}
                                startAngle={90}
                                endAngle={-270}
                                paddingAngle={1}
                                dataKey="value"
                              >
                                {act.chartData.map((entry, idx) => {
                                  let fill = entry.fill;
                                  if (isSelected) {
                                    fill = entry.name === "Actual" ? "#ffffff" : "rgba(255,255,255,0.25)";
                                  } else {
                                    if (act.name === "TOTAL" && entry.name === "Actual") {
                                      fill = "#06b6d4";
                                    }
                                  }
                                  return <Cell key={`cell-${idx}`} fill={fill} />;
                                })}
                              </Pie>
                            </PieChart>
                          </ResponsiveContainer>
                          {/* Center percentage indicator */}
                          <div className="absolute inset-0 flex items-center justify-center pointer-events-none">
                            <span className={`font-black ${
                              act.name === "TOTAL"
                                ? "text-[14px]"
                                : "text-[11.5px]"
                            } ${
                              isSelected ? "text-white" : act.name === "TOTAL" ? "text-[#154be2]" : "text-[#181a2c]"
                            }`}>
                              {act.percentage}%
                            </span>
                          </div>
                        </div>
                      </div>
                    </button>
                  );
                })}
              </div>
            </div>
          </div>

                                                  {/* Section: Conversion Sales Rate */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
            {/* Left: Conversion Sales Rate by Activity */}
            <div className="bg-white p-6 rounded-[40px] shadow-[0_12px_32px_rgba(21,75,226,0.18)] border border-[#154be2]/8 flex flex-col justify-between mt-6 lg:mt-8">
              <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-3">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="material-symbols-outlined text-primary text-sm font-semibold">
                      swap_horiz
                    </span>
                    <h3 className="text-xs font-bold text-[#181a2c] tracking-tight">
                      Conversion Sales Rate (Activity)
                    </h3>
                  </div>
                  <p className="text-[10px] text-[#8E94B7] mt-0.5">
                    Pilih filter untuk melihat impact konversi
                  </p>
                </div>

                {/* Filter Selection */}
                <select 
                  className="bg-[#f0effc] text-[#181a2c] text-[10px] font-semibold px-3 py-1.5 rounded-lg border-none outline-none cursor-pointer shrink-0"
                  value={conversionActivityFilter}
                  onChange={(e) => setConversionActivityFilter(e.target.value)}
                >
                  <option value="activity">Activity</option>
                  <option value="area">Area</option>
                  <option value="province">Province</option>
                  <option value="Sales Agronomist">Sales Agronomist</option>
                  <option value="Hybrids">Hybrids</option>
                </select>
              </div>
              <div className="flex flex-col gap-3 mt-4">
                {[
                  { name: conversionActivityFilter === "area" ? "Area 1" : conversionActivityFilter === "province" ? "Jawa Timur" : conversionActivityFilter === "Sales Agronomist" ? "Budi" : conversionActivityFilter === "Hybrids" ? "NK 212" : "Farmer Meeting", rate: conversionActivityFilter === "area" ? 82 : conversionActivityFilter === "province" ? 78 : conversionActivityFilter === "Sales Agronomist" ? 85 : conversionActivityFilter === "Hybrids" ? 90 : 68, budget: 150000, sales: 102000 },
                  { name: conversionActivityFilter === "area" ? "Area 2" : conversionActivityFilter === "province" ? "Jawa Tengah" : conversionActivityFilter === "Sales Agronomist" ? "Agus" : conversionActivityFilter === "Hybrids" ? "NK 6172" : "Demo Plot", rate: conversionActivityFilter === "area" ? 64 : conversionActivityFilter === "province" ? 65 : conversionActivityFilter === "Sales Agronomist" ? 72 : conversionActivityFilter === "Hybrids" ? 75 : 54, budget: 120000, sales: 64800 },
                  { name: conversionActivityFilter === "area" ? "Area 3" : conversionActivityFilter === "province" ? "Jawa Barat" : conversionActivityFilter === "Sales Agronomist" ? "Joko" : conversionActivityFilter === "Hybrids" ? "NK 7328" : "Field Day", rate: conversionActivityFilter === "area" ? 55 : conversionActivityFilter === "province" ? 58 : conversionActivityFilter === "Sales Agronomist" ? 60 : conversionActivityFilter === "Hybrids" ? 65 : 45, budget: 100000, sales: 45000 },
                  { name: conversionActivityFilter === "area" ? "Area 4" : conversionActivityFilter === "province" ? "Sumatera Utara" : conversionActivityFilter === "Sales Agronomist" ? "Rudi" : conversionActivityFilter === "Hybrids" ? "NK 33" : "Kiosk Visit", rate: conversionActivityFilter === "area" ? 42 : conversionActivityFilter === "province" ? 45 : conversionActivityFilter === "Sales Agronomist" ? 48 : conversionActivityFilter === "Hybrids" ? 52 : 32, budget: 200000, sales: 64000 },
                  { name: conversionActivityFilter === "area" ? "Area 5" : conversionActivityFilter === "province" ? "Sulawesi Selatan" : conversionActivityFilter === "Sales Agronomist" ? "Andi" : conversionActivityFilter === "Hybrids" ? "NK 99" : "Farmer Visit", rate: conversionActivityFilter === "area" ? 25 : conversionActivityFilter === "province" ? 28 : conversionActivityFilter === "Sales Agronomist" ? 30 : conversionActivityFilter === "Hybrids" ? 35 : 21, budget: 180000, sales: 37800 },
                ].map((item, index) => (
                  <div 
                    key={item.name} 
                    onClick={() => setActiveConversionActivity(item.name)}
                    className={`flex items-center justify-between p-3 rounded-xl border transition-all duration-200 cursor-pointer ${activeConversionActivity === item.name ? "bg-[#154be2]/5 border-[#154be2]/30 shadow-sm" : "bg-[#fbfaff] border-[#f0effc] hover:border-[#154be2]/20"}`}
                  >
                    <div className="flex items-center gap-3">
                      <div className={`size-8 shrink-0 rounded-full flex items-center justify-center font-bold text-xs ${activeConversionActivity === item.name ? "bg-[#154be2] text-white" : "bg-[#154be2]/10 text-primary"}`}>
                        #{index + 1}
                      </div>
                      <div className="min-w-0">
                        <h4 className="text-[11px] font-bold text-[#181a2c] leading-tight truncate">{item.name}</h4>
                        <div className="flex items-center gap-2 mt-0.5">
                          <p className="text-[9px] text-[#8E94B7] truncate">
                            Sales: <span className="font-semibold text-emerald-600">{Math.round(item.budget * (item.rate / 100)).toLocaleString()}</span>
                          </p>
                          <span className="text-[8px] text-slate-300">|</span>
                          <p className="text-[9px] text-[#8E94B7] truncate">
                            Budget: <span className="font-semibold text-slate-600">{item.budget.toLocaleString()}</span>
                          </p>
                        </div>
                      </div>
                    </div>
                    <div className="text-right shrink-0">
                      <p className="text-xs font-bold text-primary font-sans">{item.rate}%</p>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {/* Right: Conversion Sales Rate by Filter */}
            <div className="bg-white p-6 rounded-[40px] shadow-[0_12px_32px_rgba(21,75,226,0.18)] border border-[#154be2]/8 flex flex-col justify-between mt-6 lg:mt-8">
              <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-3">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="material-symbols-outlined text-primary text-sm font-semibold">
                      filter_alt
                    </span>
                    <h3 className="text-xs font-bold text-[#181a2c] tracking-tight">
                      Impact: {activeConversionActivity}
                    </h3>
                  </div>
                  <p className="text-[10px] text-[#8E94B7] mt-0.5">
                    Konversi penjualan berdasarkan filter
                  </p>
                </div>
                
                {/* Filter Selection */}
                <select 
                  className="bg-[#f0effc] text-[#181a2c] text-[10px] font-semibold px-3 py-1.5 rounded-lg border-none outline-none cursor-pointer"
                  value={conversionSalesFilter}
                  onChange={(e) => setConversionSalesFilter(e.target.value)}
                >
                  <option value="activity">Activity</option>
                  <option value="area">Area</option>
                  <option value="province">Province</option>
                  <option value="Sales Agronomist">Sales Agronomist</option>
                  <option value="Hybrids">Hybrids</option>
                </select>
              </div>

              <div className="flex flex-col gap-3 mt-4">
                {[
                  { name: conversionSalesFilter === "activity" ? "Farmer Meeting" : conversionSalesFilter === "area" ? "Area 1" : conversionSalesFilter === "province" ? "Jawa Timur" : conversionSalesFilter === "Sales Agronomist" ? "Budi" : "NK 212", rate: (activeConversionActivity || "").includes("1") || (activeConversionActivity || "").includes("Jawa Timur") || (activeConversionActivity || "").includes("Budi") || (activeConversionActivity || "").includes("212") || (activeConversionActivity || "").includes("Farmer Meeting") ? 75 : 45, budget: 85000, sales: 63750 },
                  { name: conversionSalesFilter === "activity" ? "Demo Plot" : conversionSalesFilter === "area" ? "Area 2" : conversionSalesFilter === "province" ? "Jawa Tengah" : conversionSalesFilter === "Sales Agronomist" ? "Agus" : "NK 6172", rate: (activeConversionActivity || "").includes("1") || (activeConversionActivity || "").includes("Jawa Timur") || (activeConversionActivity || "").includes("Budi") || (activeConversionActivity || "").includes("212") || (activeConversionActivity || "").includes("Farmer Meeting") ? 62 : 40, budget: 70000, sales: 43400 },
                  { name: conversionSalesFilter === "activity" ? "Field Day" : conversionSalesFilter === "area" ? "Area 3" : conversionSalesFilter === "province" ? "Jawa Barat" : conversionSalesFilter === "Sales Agronomist" ? "Joko" : "NK 7328", rate: (activeConversionActivity || "").includes("1") || (activeConversionActivity || "").includes("Jawa Timur") || (activeConversionActivity || "").includes("Budi") || (activeConversionActivity || "").includes("212") || (activeConversionActivity || "").includes("Farmer Meeting") ? 48 : 35, budget: 65000, sales: 31200 },
                  { name: conversionSalesFilter === "activity" ? "Kiosk Visit" : conversionSalesFilter === "area" ? "Area 4" : conversionSalesFilter === "province" ? "Sumatera Utara" : conversionSalesFilter === "Sales Agronomist" ? "Rudi" : "NK 33", rate: (activeConversionActivity || "").includes("1") || (activeConversionActivity || "").includes("Jawa Timur") || (activeConversionActivity || "").includes("Budi") || (activeConversionActivity || "").includes("212") || (activeConversionActivity || "").includes("Farmer Meeting") ? 35 : 25, budget: 50000, sales: 17500 },
                  { name: conversionSalesFilter === "activity" ? "Farmer Visit" : conversionSalesFilter === "area" ? "Area 5" : conversionSalesFilter === "province" ? "Sulawesi Selatan" : conversionSalesFilter === "Sales Agronomist" ? "Andi" : "NK 99", rate: (activeConversionActivity || "").includes("1") || (activeConversionActivity || "").includes("Jawa Timur") || (activeConversionActivity || "").includes("Budi") || (activeConversionActivity || "").includes("212") || (activeConversionActivity || "").includes("Farmer Meeting") ? 20 : 15, budget: 45000, sales: 9000 },
                ].map((item, index) => (
                  <div key={item.name} className="flex items-center justify-between p-3 rounded-xl bg-[#fbfaff] border border-[#f0effc] hover:border-[#154be2]/20 transition-all duration-200">
                    <div className="flex items-center gap-3">
                      <div className="size-8 shrink-0 rounded-full flex items-center justify-center font-bold text-xs bg-[#154be2]/10 text-primary">
                        #{index + 1}
                      </div>
                      <div className="min-w-0">
                        <h4 className="text-[11px] font-bold text-[#181a2c] leading-tight truncate">{item.name}</h4>
                        <div className="flex items-center gap-2 mt-0.5">
                          <p className="text-[9px] text-[#8E94B7] truncate">
                            Sales: <span className="font-semibold text-emerald-600">{Math.round(item.budget * (item.rate / 100)).toLocaleString()}</span>
                          </p>
                          <span className="text-[8px] text-slate-300">|</span>
                          <p className="text-[9px] text-[#8E94B7] truncate">
                            Budget: <span className="font-semibold text-slate-600">{item.budget.toLocaleString()}</span>
                          </p>
                        </div>
                      </div>
                    </div>
                    <div className="text-right shrink-0">
                      <p className="text-xs font-bold text-primary font-sans">{item.rate}%</p>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
          
          {/* Section: History Bulanan */}
          <div className="bg-white p-6 rounded-[40px] shadow-[0_12px_32px_rgba(21,75,226,0.18)] border border-[#154be2]/8 flex flex-col justify-between mt-6">
            <div className="flex flex-col md:flex-row md:items-center md:justify-between mb-6 pb-4 border-b border-[#f0effc]/60 gap-4">
              <div>
                <div className="flex items-center gap-2">
                  <span className="material-symbols-outlined text-primary text-sm font-semibold border-b border-primary/20 pb-0.5">
                    timeline
                  </span>
                  <h3 className="text-xs font-bold text-[#181a2c] tracking-tight">
                    Activity Performance Trend
                  </h3>
                </div>
                <p className="text-[10px] text-[#8E94B7] mt-0.5">
                  Analisis perbandingan histori total volume perkembangan data
                  dari bulan ke bulan
                </p>
              </div>

              {/* Metric Toggle for Trend Chart */}
              <div className="flex items-center gap-2 self-start md:self-auto bg-[#fbfaff] px-3.5 py-1.5 rounded-xl border border-[#e2e8f0]/40 shrink-0">
                <div className="flex items-center gap-4 select-none">
                  <button
                    type="button"
                    onClick={() => setShowBudgetBar(prev => !prev)}
                    className={`flex items-center gap-2 hover:opacity-85 transition-all cursor-pointer ${!showBudgetBar ? "opacity-35 line-through" : ""}`}
                    title="Klik untuk menyembunyikan/menampilkan Budget"
                  >
                    <div className="w-3.5 h-3.5 rounded-sm bg-gradient-to-b from-[#ea580c] to-[#c2410c] shadow-sm"></div>
                    <span className="text-[10.5px] font-bold text-slate-700 uppercase tracking-wide">Budget</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => setShowActualBar(prev => !prev)}
                    className={`flex items-center gap-2 hover:opacity-85 transition-all cursor-pointer ${!showActualBar ? "opacity-35 line-through" : ""}`}
                    title="Klik untuk menyembunyikan/menampilkan Actual"
                  >
                    <div className="w-3.5 h-3.5 rounded-sm bg-gradient-to-b from-[#f97316] to-[#ea580c] shadow-sm"></div>
                    <span className="text-[10.5px] font-bold text-slate-700 uppercase tracking-wide">Actual</span>
                  </button>
                </div>
              </div>
            </div>

            <div className="h-64 w-full font-sans">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart
                  data={overviewHistoryData}
                  margin={{ top: 38, right: 10, left: -10, bottom: 0 }}
                  barGap={currentBarGap}
                  barCategoryGap={currentBarCategoryGap}
                  onMouseMove={(state) => {
                    if (state && state.activeLabel) {
                      setHoveredLabel(state.activeLabel);
                      if (state.activeLabel !== dismissedTooltipLabel) {
                        setDismissedTooltipLabel(null);
                      }
                    } else {
                      setHoveredLabel(null);
                    }
                  }}
                  onMouseLeave={() => {
                    setHoveredLabel(null);
                  }}
                >
                  <defs>
                    <linearGradient
                      id="colorTrendBudget"
                      x1="0"
                      y1="0"
                      x2="0"
                      y2="1"
                    >
                      <stop
                        offset="0%"
                        stopColor="#154be2"
                        stopOpacity={0.95}
                      />
                      <stop
                        offset="100%"
                        stopColor="#3b82f6"
                        stopOpacity={0.7}
                      />
                    </linearGradient>
                    <linearGradient
                      id="colorTrendActual"
                      x1="0"
                      y1="0"
                      x2="0"
                      y2="1"
                    >
                      <stop
                        offset="0%"
                        stopColor="#06b6d4"
                        stopOpacity={1.0}
                      />
                      <stop
                        offset="100%"
                        stopColor="#22d3ee"
                        stopOpacity={1.0}
                      />
                    </linearGradient>
                  </defs>
                  <CartesianGrid
                    strokeDasharray="4 4"
                    vertical={false}
                    stroke="#e2e8f0"
                  />
                  <XAxis
                    dataKey="name"
                    tick={<CustomXAxisTick chartData={overviewHistoryData} metricType={overviewMetricFilter} />}
                    axisLine={false}
                    tickLine={false}
                    interval={0}
                    height={65}
                  />
                  <YAxis
                    hide={true}
                    domain={[0, (dataMax) => (dataMax === 0 ? 100 : Math.round(dataMax * 1.25))]}
                    tick={{ fill: "#8E94B7", fontSize: 9, fontWeight: 500 }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <Tooltip
                    cursor={{ fill: "rgba(21, 75, 226, 0.03)" }}
                    content={<CustomChartTooltip metricType={overviewMetricFilter} dismissedLabel={dismissedTooltipLabel} />}
                  />
                  <Bar
                    hide={!showBudgetBar}
                    dataKey={
                      overviewMetricFilter === "activity"
                        ? "budgetActivity"
                        : overviewMetricFilter === "reach"
                        ? "budgetReach"
                        : "budgetNominal"
                    }
                    name="Budget"
                    fill="url(#colorTrendBudget)"
                    radius={[6, 6, 0, 0]}
                    maxBarSize={currentMaxBarSize}
                    background={<CustomBarBackground data={overviewHistoryData} activeKey={activeMainBarKey} />}
                  >
                    {overviewHistoryData.map((entry, index) => {
                      const isActive = activeMainBarKey === entry.name;
                      return (
                        <Cell
                          key={`cell-trend-budget-${index}`}
                          cursor="pointer"
                          fill={isActive ? "#ea580c" : "url(#colorTrendBudget)"}
                          fillOpacity={1.0}
                          onClick={() => {
                            clickedBarRef.current = true;
                            if (activeMainBarKey === entry.name) {
                              setDismissedTooltipLabel(entry.name);
                            } else {
                              setDismissedTooltipLabel(null);
                            }
                            setActiveMainBarKey(prev => prev === entry.name ? null : entry.name);
                          }}
                        />
                      );
                    })}
                    <LabelList
                      dataKey={
                        overviewMetricFilter === "activity"
                          ? "budgetActivity"
                          : overviewMetricFilter === "reach"
                          ? "budgetReach"
                          : "budgetNominal"
                      }
                      content={<CustomBudgetLabel metricType={overviewMetricFilter} />}
                    />
                  </Bar>
                  <Bar
                    hide={!showActualBar}
                    dataKey={
                      overviewMetricFilter === "activity"
                        ? "actualActivity"
                        : overviewMetricFilter === "reach"
                        ? "actualReach"
                        : "actualNominal"
                    }
                    name="Actual"
                    fill="url(#colorTrendActual)"
                    radius={[6, 6, 0, 0]}
                    maxBarSize={currentMaxBarSize}
                    background={<CustomBarBackground data={overviewHistoryData} activeKey={activeMainBarKey} />}
                  >
                    {overviewHistoryData.map((entry, index) => {
                      const isActive = activeMainBarKey === entry.name;
                      return (
                        <Cell
                          key={`cell-trend-actual-${index}`}
                          cursor="pointer"
                          fill={isActive ? "#f97316" : "url(#colorTrendActual)"}
                          fillOpacity={1.0}
                          onClick={() => {
                            clickedBarRef.current = true;
                            if (activeMainBarKey === entry.name) {
                              setDismissedTooltipLabel(entry.name);
                            } else {
                              setDismissedTooltipLabel(null);
                            }
                            setActiveMainBarKey(prev => prev === entry.name ? null : entry.name);
                          }}
                        />
                      );
                    })}
                    <LabelList
                      dataKey={
                        overviewMetricFilter === "activity"
                          ? "actualActivity"
                          : overviewMetricFilter === "reach"
                          ? "actualReach"
                          : "actualNominal"
                      }
                      content={<CustomActualLabel metricType={overviewMetricFilter} />}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Focused Chart Modal */}
          {isChartFocusedModalOpen && (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
              <div className="relative w-full max-w-5xl bg-white rounded-[28px] shadow-[0_24px_64px_rgba(21,75,226,0.15)] border border-[#154be2]/10 p-6 md:p-8 flex flex-col justify-between max-h-[90vh] overflow-hidden">
                {/* Header */}
                <div className="flex items-start justify-between pb-4 border-b border-[#f0effc]">
                  <div>
                    <div className="flex items-center gap-2 text-primary font-bold">
                      <span className="material-symbols-outlined text-[#154be2]">
                        {focusedChartType === "sub" ? "bar_chart" : "analytics"}
                      </span>
                      <span className="text-[14px] uppercase tracking-wider text-[#154be2] font-black">
                        {focusedChartType === "sub" ? "Detail Fokus Grafik Sub" : "Detail Fokus Grafik Utama"}
                      </span>
                    </div>
                    <h3 className="text-lg font-black text-[#181a2c] tracking-tight mt-1">
                      {focusedChartType === "sub" ? (
                        subGroupDimension === "area"
                          ? "Sub Budget Effectiveness Wilayah (Area)"
                          : subGroupDimension === "province"
                            ? "Sub Budget Effectiveness per Provinsi"
                            : subGroupDimension === "sales_agronomist"
                              ? "Sub Budget Effectiveness Sales Agronomist (SA)"
                              : subGroupDimension === "hybrid" || subGroupDimension === "material"
                                ? "Sub Budget Effectiveness per Hybrid"
                                : "Sub Budget Effectiveness per Activity"
                      ) : (
                        overviewGroupDimension === "area"
                          ? "Budget Effectiveness Wilayah (Area)"
                          : overviewGroupDimension === "province"
                            ? "Budget Effectiveness per Provinsi"
                            : overviewGroupDimension === "sales_agronomist"
                              ? "Budget Effectiveness Sales Agronomist (SA)"
                              : overviewGroupDimension === "hybrid" || overviewGroupDimension === "material"
                                ? "Budget Effectiveness per Hybrid"
                                : "Budget Effectiveness per Activity"
                      )}
                    </h3>
                  </div>
                  <button
                    onClick={() => setIsChartFocusedModalOpen(false)}
                    className="p-2 rounded-full hover:bg-slate-100 transition-colors text-slate-400 hover:text-slate-700"
                  >
                    <span className="material-symbols-outlined text-xl">close</span>
                  </button>
                </div>

                {/* Controls inside Modal */}
                <div className="flex flex-wrap items-center justify-between gap-4 py-4 bg-[#fbfaff] px-4 rounded-2xl border border-[#e2e8f0]/60 my-4 shrink-0">
                  <div className="flex flex-wrap items-center gap-4">
                    <div className="flex items-center gap-2">
                      <span className="text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider">Dimensi:</span>
                      {focusedChartType === "sub" ? (
                        <select
                          value={subGroupDimension}
                          onChange={(e: any) => setSubGroupDimension(e.target.value as any)}
                          className="bg-white border border-[#e2e8f0] rounded-xl px-2.5 py-1 text-xs font-black text-[#154be2] focus:outline-none focus:ring-1 focus:ring-[#154be2] cursor-pointer"
                        >
                          <option value="area">Area</option>
                          <option value="province">Province</option>
                          <option value="sales_agronomist">Sales Agronomist</option>
                          <option value="hybrid">Hybrids</option>
                          <option value="activity">Activity</option>
                        </select>
                      ) : (
                        <select
                          value={overviewGroupDimension}
                          onChange={(e: any) => setOverviewGroupDimension(e.target.value as any)}
                          className="bg-white border border-[#e2e8f0] rounded-xl px-2.5 py-1 text-xs font-black text-[#154be2] focus:outline-none focus:ring-1 focus:ring-[#154be2] cursor-pointer"
                        >
                          <option value="area">Area</option>
                          <option value="province">Province</option>
                          <option value="sales_agronomist">Sales Agronomist</option>
                          <option value="hybrid">Hybrids</option>
                          <option value="activity">Activity</option>
                        </select>
                      )}
                    </div>
                    <div className="flex items-center gap-2">
                      <div className="flex bg-slate-100 p-0.5 rounded-xl border border-slate-200">
                        <button
                          onClick={() => setOverviewMetricFilter("activity")}
                          className={`px-5 py-2 rounded-lg text-xs font-bold transition-all duration-200 ${
                            overviewMetricFilter === "activity"
                              ? "bg-[#154be2] text-white shadow-sm"
                              : "text-slate-600 hover:text-slate-900"
                          }`}
                        >
                          Activity
                        </button>
                        <button
                          onClick={() => setOverviewMetricFilter("nominal")}
                          className={`px-5 py-2 rounded-lg text-xs font-bold transition-all duration-200 ${
                            overviewMetricFilter === "nominal"
                              ? "bg-[#154be2] text-white shadow-sm"
                              : "text-slate-600 hover:text-slate-900"
                          }`}
                        >
                          Nominal
                        </button>
                        <button
                          onClick={() => setOverviewMetricFilter("reach")}
                          className={`px-5 py-2 rounded-lg text-xs font-bold transition-all duration-200 ${
                            overviewMetricFilter === "reach"
                              ? "bg-[#154be2] text-white shadow-sm"
                              : "text-slate-600 hover:text-slate-900"
                          }`}
                        >
                          Reach
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Legend */}
                  <div className="flex items-center gap-4 ml-auto select-none">
                    <button
                      type="button"
                      onClick={() => setShowBudgetBar(prev => !prev)}
                      className={`flex items-center gap-2 hover:opacity-85 transition-all cursor-pointer ${!showBudgetBar ? "opacity-35 line-through" : ""}`}
                      title="Klik untuk menyembunyikan/menampilkan Budget"
                    >
                      <span className="size-4 rounded-[4px] bg-gradient-to-tr from-[#154be2] to-[#3b82f6]" />
                      <span className="text-sm font-extrabold text-[#4e5572]">
                        Budget
                      </span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setShowActualBar(prev => !prev)}
                      className={`flex items-center gap-2 hover:opacity-85 transition-all cursor-pointer ${!showActualBar ? "opacity-35 line-through" : ""}`}
                      title="Klik untuk menyembunyikan/menampilkan Actual"
                    >
                      <span className="size-4 rounded-[4px] bg-gradient-to-tr from-[#06b6d4] to-[#22d3ee]" />
                      <span className="text-sm font-extrabold text-[#4e5572]">
                        Actual
                      </span>
                    </button>
                  </div>
                </div>

                {/* Chart container in Modal */}
                <div className="w-full flex-1 min-h-[350px] overflow-x-auto overflow-y-hidden scrollbar-thin select-none">
                  <div
                    style={{
                      minWidth: `${Math.max(
                        800,
                        (focusedChartType === "sub"
                          ? overviewStats.subChartData?.length || 0
                          : overviewStats.areaChartData?.length || 0) * 120
                      )}px`,
                      width: "100%",
                      height: "350px",
                    }}
                    className="font-sans"
                  >
                    <ResponsiveContainer width="100%" height="100%">
                       <BarChart
                        data={focusedChartType === "sub" ? overviewStats.subChartData : overviewStats.areaChartData}
                        margin={{ top: 45, right: 15, left: -10, bottom: 35 }}
                        barGap={currentBarGap}
                        barCategoryGap={currentBarCategoryGap}
                        onMouseMove={(state) => {
                          if (state && state.activeLabel) {
                            setHoveredLabel(state.activeLabel);
                            if (focusedChartType === "sub") {
                              if (state.activeLabel !== dismissedSubTooltipLabel) {
                                setDismissedSubTooltipLabel(null);
                              }
                            } else {
                              if (state.activeLabel !== dismissedTooltipLabel) {
                                setDismissedTooltipLabel(null);
                              }
                            }
                          } else {
                            setHoveredLabel(null);
                          }
                        }}
                        onMouseLeave={() => {
                          setHoveredLabel(null);
                        }}
                      >
                        <defs>
                          <linearGradient id="modalColorAreaPog" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stopColor="#154be2" stopOpacity={0.95} />
                            <stop offset="100%" stopColor="#3b82f6" stopOpacity={0.7} />
                          </linearGradient>
                          <linearGradient id="modalColorAreaStock" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stopColor="#06b6d4" stopOpacity={1.0} />
                            <stop offset="100%" stopColor="#22d3ee" stopOpacity={1.0} />
                          </linearGradient>
                        </defs>
                        <CartesianGrid strokeDasharray="4 4" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="name" tick={<CustomXAxisTick chartData={focusedChartType === "sub" ? overviewStats.subChartData : overviewStats.areaChartData} metricType={overviewMetricFilter} />} axisLine={false} tickLine={false} interval={0} height={65} />
                        <YAxis hide={true} domain={[0, (dataMax: any) => (dataMax === 0 ? 100 : Math.round(dataMax * 1.25))]} tick={{ fill: "#8E94B7", fontSize: 10, fontWeight: 500 }} axisLine={false} tickLine={false} />
                        <Tooltip
                          cursor={{ fill: "rgba(21, 75, 226, 0.03)" }}
                          content={<CustomChartTooltip metricType={overviewMetricFilter} dismissedLabel={focusedChartType === "sub" ? dismissedSubTooltipLabel : dismissedTooltipLabel} />}
                        />
                        <Bar
                          hide={!showBudgetBar}
                          dataKey={
                            overviewMetricFilter === "activity"
                              ? "budgetActivity"
                              : overviewMetricFilter === "reach"
                              ? "budgetReach"
                              : "budgetNominal"
                          }
                          name="Budget"
                          fill="url(#modalColorAreaPog)"
                          radius={[6, 6, 0, 0]}
                          maxBarSize={currentMaxBarSize}
                          background={<CustomBarBackground data={focusedChartType === "sub" ? overviewStats.subChartData : overviewStats.areaChartData} activeKey={focusedChartType === "sub" ? activeSubBarKey : activeMainBarKey} />}
                        >
                          {(focusedChartType === "sub" ? overviewStats.subChartData : overviewStats.areaChartData)?.map((entry: any, index: number) => {
                            const isMain = focusedChartType === "main";
                            const isActive = (focusedChartType === "sub" ? activeSubBarKey : activeMainBarKey) === entry.name;
                            return (
                              <Cell
                                key={`modal-cell-budget-${index}`}
                                cursor="pointer"
                                fill={isActive ? "#ea580c" : "url(#modalColorAreaPog)"}
                                fillOpacity={1.0}
                                onClick={() => {
                                  if (isMain) {
                                    clickedBarRef.current = true;
                                    if (activeMainBarKey === entry.name) {
                                      setDismissedTooltipLabel(entry.name);
                                    } else {
                                      setDismissedTooltipLabel(null);
                                    }
                                    setActiveMainBarKey(prev => prev === entry.name ? null : entry.name);
                                  } else {
                                    clickedBarRef.current = true;
                                    if (activeSubBarKey === entry.name) {
                                      setDismissedSubTooltipLabel(entry.name);
                                    } else {
                                      setDismissedSubTooltipLabel(null);
                                    }
                                    setActiveSubBarKey(prev => prev === entry.name ? null : entry.name);
                                  }
                                }}
                              />
                            );
                          })}
                          <LabelList
                            dataKey={
                              overviewMetricFilter === "activity"
                                ? "budgetActivity"
                                : overviewMetricFilter === "reach"
                                ? "budgetReach"
                                : "budgetNominal"
                            }
                            content={<CustomBudgetLabel metricType={overviewMetricFilter} />}
                          />
                        </Bar>
                        <Bar
                          hide={!showActualBar}
                          dataKey={
                            overviewMetricFilter === "activity"
                              ? "actualActivity"
                              : overviewMetricFilter === "reach"
                              ? "actualReach"
                              : "actualNominal"
                          }
                          name="Actual"
                          fill="url(#modalColorAreaStock)"
                          radius={[6, 6, 0, 0]}
                          maxBarSize={currentMaxBarSize}
                          background={<CustomBarBackground data={focusedChartType === "sub" ? overviewStats.subChartData : overviewStats.areaChartData} activeKey={focusedChartType === "sub" ? activeSubBarKey : activeMainBarKey} />}
                        >
                          {(focusedChartType === "sub" ? overviewStats.subChartData : overviewStats.areaChartData)?.map((entry: any, index: number) => {
                            const isMain = focusedChartType === "main";
                            const isActive = (focusedChartType === "sub" ? activeSubBarKey : activeMainBarKey) === entry.name;
                            return (
                              <Cell
                                key={`modal-cell-actual-${index}`}
                                cursor="pointer"
                                fill={isActive ? "#f97316" : "url(#modalColorAreaStock)"}
                                fillOpacity={1.0}
                                onClick={() => {
                                  if (isMain) {
                                    clickedBarRef.current = true;
                                    if (activeMainBarKey === entry.name) {
                                      setDismissedTooltipLabel(entry.name);
                                    } else {
                                      setDismissedTooltipLabel(null);
                                    }
                                    setActiveMainBarKey(prev => prev === entry.name ? null : entry.name);
                                  } else {
                                    clickedBarRef.current = true;
                                    if (activeSubBarKey === entry.name) {
                                      setDismissedSubTooltipLabel(entry.name);
                                    } else {
                                      setDismissedSubTooltipLabel(null);
                                    }
                                    setActiveSubBarKey(prev => prev === entry.name ? null : entry.name);
                                  }
                                }}
                              />
                            );
                          })}
                          <LabelList
                            dataKey={
                              overviewMetricFilter === "activity"
                                ? "actualActivity"
                                : overviewMetricFilter === "reach"
                                ? "actualReach"
                                : "actualNominal"
                            }
                            content={<CustomActualLabel metricType={overviewMetricFilter} />}
                          />
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {activeTab === "home" && (
        <div className="animate-in fade-in slide-in-from-left-4 duration-300">
          {isBusinessAnalyst ? (
            <div className="bg-white p-6 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.12)] border border-[#edecff] text-center max-w-md mx-auto my-12">
              <span className="material-symbols-outlined text-[48px] text-primary/40 mb-3">
                analytics
              </span>
              <h2 className="text-base font-semibold text-[#181a2c] mb-1.5">
                Akses Analis Bisnis
              </h2>
              <p className="text-xs text-[#8E94B7] leading-relaxed">
                Sebagai Business Analyst, Anda tidak perlu menginput stok secara
                manual untuk masing-masing Toko/Channel Partner melainkan
                memantau visualisasi, performa, dan ringkasan data. Silakan buka
                tab <strong className="text-primary font-medium">Stock</strong>{" "}
                atau <strong className="text-primary font-medium">POG</strong>.
              </p>
            </div>
          ) : (
            <>
              <div className="mb-4 ml-1">
                <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
                  Channel{" "}
                  <span className="text-primary font-bold">Partner</span>
                </h1>
              </div>
              <div className="relative mb-8" ref={dropdownRef}>
                <button
                  onClick={() => setIsDropdownOpen(!isDropdownOpen)}
                  className="flex items-center justify-between w-full px-6 py-4 bg-gradient-to-r from-primary to-cyan-400 text-white rounded-full shadow-[0_12px_32px_rgba(21,75,226,0.25)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.35)] hover:scale-[1.01] active:scale-[0.99] transition-all relative z-30"
                >
                  <span className="font-semibold text-sm truncate">
                    {selectedKiosk}
                  </span>
                  <span className="material-symbols-outlined text-white/80">
                    unfold_more
                  </span>
                </button>
                {isDropdownOpen && (
                  <div className="absolute left-0 right-0 mt-2 bg-white/95 backdrop-blur-md shadow-[0_12px_32px_rgba(21,75,226,0.15)] z-[100] rounded-[24px] p-4 animate-in fade-in slide-in-from-top-4">
                    <input
                      type="text"
                      placeholder="Search channel..."
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      className="w-full h-11 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.06)] rounded-full px-5 text-xs font-semibold outline-none focus:ring-1 focus:ring-primary/20 mb-3"
                    />
                    <div className="max-h-[200px] overflow-y-auto custom-scrollbar space-y-1">
                      {filteredKiosks.map((k, i) => (
                        <button
                          key={i}
                          onClick={() => {
                            setSelectedKiosk(k.name);
                            setIsDropdownOpen(false);
                            setSearchTerm("");
                          }}
                          className={`w-full text-left px-5 py-2.5 rounded-full font-semibold text-[11px] uppercase transition-all ${selectedKiosk === k.name ? "bg-[#edecff] text-primary" : "text-[#635b6e] hover:bg-[#fbf8ff]"}`}
                        >
                          {k.name}
                        </button>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              <div className="mb-4 ml-1">
                <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
                  Tambah <span className="text-primary font-bold">LOT</span>
                </h1>
              </div>
              <div className="bg-white p-5 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.15)] mb-6">
                <div className="space-y-4">
                  <div className="grid grid-cols-5 gap-3">
                    <div className="col-span-3 relative">
                      <button
                        type="button"
                        onClick={() => !isLotChecking && setIsScannerOpen(true)}
                        className={`absolute left-3 top-1/2 -translate-y-1/2 size-9 rounded-full flex items-center justify-center transition-all z-20 ${
                          isLotChecking
                            ? "cursor-not-allowed text-primary animate-spin"
                            : "cursor-pointer hover:bg-[#f4f2ff] text-primary active:scale-[0.93]"
                        }`}
                        title="Scan QR / Barcode"
                      >
                        <span className="material-symbols-outlined text-lg leading-none">
                          {isLotChecking ? "sync" : "photo_camera"}
                        </span>
                      </button>
                      <input
                        value={lotNo}
                        onChange={(e) => setLotNo(e.target.value)}
                        className="w-full h-14 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full pl-14 pr-6 font-semibold text-xs outline-none focus:ring-1 focus:ring-primary/20 transition-all"
                        placeholder="Batch / Lot No"
                      />
                    </div>
                    <input
                      value={qty}
                      onChange={(e) => setQty(e.target.value)}
                      className="col-span-2 h-14 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-6 font-semibold text-xs outline-none focus:ring-1 focus:ring-primary/20 transition-all"
                      placeholder="Qty (Kg)"
                      type="number"
                    />
                  </div>

                  {isLotChecking && (
                    <div className="bg-[#f0f3ff] border border-[#dce2ff] p-3.5 rounded-[18px] flex items-center gap-2.5 animate-pulse shadow-sm">
                      <span className="material-symbols-outlined text-primary text-lg animate-spin">
                        sync
                      </span>
                      <p className="text-xs font-semibold text-primary">
                        Mengecek LOT di database...
                      </p>
                    </div>
                  )}

                  {!isLotChecking &&
                    lotIntel &&
                    typeof lotIntel === "object" && (
                      <div className="space-y-2">
                        <div className="bg-emerald-50 border border-emerald-100 p-3.5 rounded-[18px] flex items-center gap-2.5 animate-in slide-in-from-top-2 shadow-sm">
                          <span className="material-symbols-outlined text-emerald-500 text-lg">
                            check_circle
                          </span>
                          <p className="text-xs font-semibold text-emerald-800">
                            ✅ LOT ditemukan di database!
                          </p>
                        </div>
                        <div className="bg-white shadow-[0_4px_16px_rgba(21,75,226,0.06)] p-4 rounded-[20px] animate-in slide-in-from-top-2">
                          <div className="mb-3">
                            <p className="text-[8.5px] font-bold text-[#8E94B7] uppercase tracking-wider mb-0.5">
                              Hybrid Description
                            </p>
                            <p className="text-xs font-semibold text-[#181a2c] leading-snug">
                              {lotIntel.desc}
                            </p>
                          </div>
                          <div className="flex items-center gap-4 border-t border-[#edecff] pt-3">
                            <div className="flex-1">
                              <p className="text-[8.5px] font-bold text-[#8E94B7] uppercase tracking-wider mb-0.5">
                                Dr Date
                              </p>
                              <p className="text-[10px] font-semibold text-[#181a2c]">
                                {lotIntel.drDate}
                              </p>
                            </div>
                            <div className="w-px h-6 bg-[#edecff]" />
                            <div className="flex-1">
                              <p className="text-[8.5px] font-bold text-red-400 uppercase tracking-wider mb-0.5">
                                Exp Date
                              </p>
                              <p className="text-[10px] font-bold text-red-700">
                                {lotIntel.expDate}
                              </p>
                            </div>
                          </div>
                        </div>
                      </div>
                    )}

                  {!isLotChecking &&
                    lotIntel &&
                    typeof lotIntel === "string" && (
                      <div className="space-y-2">
                        <div className="bg-emerald-50 border border-emerald-100 p-3.5 rounded-[18px] flex items-center gap-2.5 animate-in slide-in-from-top-2 shadow-sm">
                          <span className="material-symbols-outlined text-emerald-500 text-lg">
                            check_circle
                          </span>
                          <p className="text-xs font-semibold text-emerald-800">
                            ✅ LOT ditemukan di database!
                          </p>
                        </div>
                        <div className="bg-white shadow-[0_4px_16px_rgba(21,75,226,0.06)] p-4 rounded-[18px] animate-in slide-in-from-top-2">
                          <p className="text-[8.5px] font-bold text-primary uppercase mb-0.5">
                            Lot Metadata Note
                          </p>
                          <p className="text-xs font-semibold text-[#181a2c]">
                            {lotIntel}
                          </p>
                        </div>
                      </div>
                    )}

                  {!isLotChecking && isLotNotFound && (
                    <div className="space-y-4 animate-in fade-in duration-200">
                      {/* Warning Message */}
                      <div className="bg-red-50 border border-red-100 p-3.5 rounded-[18px] flex items-center gap-2.5 shadow-sm">
                        <span className="material-symbols-outlined text-red-500 text-lg flex-shrink-0">
                          warning
                        </span>
                        <div className="flex flex-col">
                          <p className="text-xs font-bold text-red-700 leading-tight">
                            ⚠️ LOT tidak ditemukan di database!
                          </p>
                          <p className="text-[10px] text-red-600 mt-0.5 font-medium leading-relaxed">
                            LOT ini belum terdaftar. Agar data stock tetap valid, Anda wajib melengkapi data Hybrid & Komoditas secara manual di bawah ini.
                          </p>
                        </div>
                      </div>

                      {/* 1. Hybrid/Varietas Input */}
                      <div className="space-y-1">
                        <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 block">
                          Hybrid / Varietas <span className="text-red-500">*</span>
                        </label>
                        <div className="relative">
                          <select
                            value={manualHybrid}
                            onChange={(e) => setManualHybrid(e.target.value)}
                            className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all appearance-none pr-10"
                            required
                          >
                            <option value="">-- Pilih Hybrid / Varietas --</option>
                            {filterOptions.materials.map((mat) => (
                              <option key={mat} value={mat}>
                                {mat}
                              </option>
                            ))}
                            <option value="CUSTOM">-- Ketik Manual (Varietas Baru) --</option>
                          </select>
                          <span className="material-symbols-outlined absolute right-4 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none text-lg">
                            expand_more
                          </span>
                        </div>
                        {manualHybrid === "CUSTOM" && (
                          <input
                            type="text"
                            value={manualHybridCustom}
                            onChange={(e) => setManualHybridCustom(e.target.value)}
                            placeholder="Ketik Nama Hybrid Baru..."
                            className="w-full h-11 mt-2 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all"
                            required
                          />
                        )}
                      </div>

                      {/* 2. Commodity / Crop Input */}
                      <div className="space-y-1">
                        <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 block">
                          Komoditas <span className="text-red-500">*</span>
                        </label>
                        <div className="relative">
                          <select
                            value={manualCrop}
                            onChange={(e) => setManualCrop(e.target.value)}
                            className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all appearance-none pr-10"
                            required
                          >
                            {["Field Corn", "Fresh Corn", "Vegetables"].map((crop) => (
                              <option key={crop} value={crop}>
                                {crop}
                              </option>
                            ))}
                          </select>
                          <span className="material-symbols-outlined absolute right-4 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none text-lg">
                            expand_more
                          </span>
                        </div>
                      </div>

                      {/* 3. Dates Grid */}
                      <div className="grid grid-cols-2 gap-3">
                        {/* Shipping Date */}
                        <div className="space-y-1">
                          <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 block">
                            Tgl DR / Shipping
                          </label>
                          <input
                            type="date"
                            value={manualDrDate}
                            onChange={(e) => setManualDrDate(e.target.value)}
                            className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all"
                          />
                        </div>

                        {/* Expired Date */}
                        <div className="space-y-1">
                          <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 block">
                            Expired Date
                          </label>
                          <input
                            type="date"
                            value={manualExpDate}
                            onChange={(e) => setManualExpDate(e.target.value)}
                            className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all"
                          />
                        </div>
                      </div>
                    </div>
                  )}

                  <button
                    onClick={handleAddLocal}
                    disabled={
                      !lotNo ||
                      !qty ||
                      (isLotNotFound &&
                        (!manualHybrid ||
                          (manualHybrid === "CUSTOM" && !manualHybridCustom.trim()) ||
                          !manualCrop))
                    }
                    className={`w-full h-14 rounded-full font-semibold text-xs uppercase tracking-wider flex items-center justify-center gap-2 transition-all active:scale-[0.98] ${
                      !lotNo ||
                      !qty ||
                      (isLotNotFound &&
                        (!manualHybrid ||
                          (manualHybrid === "CUSTOM" && !manualHybridCustom.trim()) ||
                          !manualCrop))
                        ? "bg-[#e0e0fa] text-[#8E94B7] shadow-none cursor-not-allowed"
                        : "bg-gradient-to-r from-emerald-500 to-[#00D2FF] text-white shadow-[0_8px_20px_rgba(16,185,129,0.2)] hover:opacity-95"
                    }`}
                  >
                    <span className="material-symbols-outlined text-[18px]">
                      add_box
                    </span>{" "}
                    Tambah ke List
                  </button>
                </div>
              </div>

              <DetailItemSection
                items={workingData.filter(
                  (item) => item.kiosk === selectedKiosk,
                )}
                onEdit={(item) => setEditModal({ isOpen: true, item })}
                onDelete={(item) => setDeleteModal({ isOpen: true, item })}
                onUploadActivity={handleUploadActivity}
                isSyncing={isSyncing}
                hasChanges={hasChanges}
                title="Existing Stock"
                subtitle={selectedKiosk}
                category={
                  kiosks.find(
                    (k) =>
                      cleanForMatch(k.name) === cleanForMatch(selectedKiosk),
                  )?.category || "Uncategorized"
                }
              />

              <EditModal
                isOpen={editModal.isOpen}
                item={editModal.item}
                onClose={() => setEditModal({ isOpen: false, item: null })}
                onSave={handleEditLocal}
                isSaving={isActionLoading}
                allHybrids={filterOptions.materials}
              />
              <ConfirmModal
                isOpen={deleteModal.isOpen}
                onClose={() => setDeleteModal({ isOpen: false, item: null })}
                onConfirm={handleDeleteLocal}
                isProcessing={isActionLoading}
              />
              <QrScanModal
                isOpen={isScannerOpen}
                onClose={() => setIsScannerOpen(false)}
                onScanSuccess={(val) => setLotNo(val)}
              />
            </>
          )}
        </div>
      )}

      {activeTab === "partner" && (
        <div className="animate-in fade-in slide-in-from-right-4 duration-300">
          <div className="mb-6 ml-1 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
            <div>
              <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
                Mapping <span className="text-primary font-bold">Partner</span>
              </h1>
              <p className="text-[11px] text-[#8E94B7] font-semibold uppercase tracking-wider mt-0.5">
                Kelola Channel Area & Tim
              </p>
            </div>

            {/* Sub-navigation Menu */}
            <div className="flex gap-1 p-1 bg-[#edecff]/45 rounded-full w-full md:w-auto max-w-xs md:max-w-none border border-[#edecff]/60 shadow-[inset_0_1px_2px_rgba(21,75,226,0.03)] shrink-0 self-start md:self-auto">
              <button
                  onClick={() => setPartnerSubTab("team")}
                  className={`flex-1 md:flex-initial px-4 py-1.5 rounded-full font-bold text-[9.5px] uppercase tracking-wider transition-all duration-200 cursor-pointer ${
                    partnerSubTab === "team"
                      ? "bg-gradient-to-r from-primary to-cyan-400 text-white shadow-[0_4px_10px_rgba(21,75,226,0.18)]"
                      : "text-[#8E94B7] hover:bg-white/50 hover:text-[#181a2c]"
                  }`}
                >
                  <div className="flex items-center justify-center gap-1.5">
                    <span className="material-symbols-outlined text-[13px] leading-none">
                      groups
                    </span>
                    <span>Tim & Hirarki</span>
                  </div>
                </button>
                <button
                  onClick={() => setPartnerSubTab("channel")}
                  className={`flex-1 md:flex-initial px-4 py-1.5 rounded-full font-bold text-[9.5px] uppercase tracking-wider transition-all duration-200 cursor-pointer ${
                    partnerSubTab === "channel"
                      ? "bg-gradient-to-r from-primary to-cyan-400 text-white shadow-[0_4px_10px_rgba(21,75,226,0.18)]"
                      : "text-[#8E94B7] hover:bg-white/50 hover:text-[#181a2c]"
                  }`}
                >
                  <div className="flex items-center justify-center gap-1.5">
                    <span className="material-symbols-outlined text-[13px] leading-none">
                      storefront
                    </span>
                    <span>Mapping Channel</span>
                  </div>
                </button>
              </div>
          </div>

          {partnerSubTab === "team" &&
            (() => {
              const getDirectSubordinates = (parentName: string) => {
                return teamMembers.filter((member) => {
                  if (matchNames(member, parentName)) return false;
                  if (matchNames(member, userData.name)) return false;

                  const mPos = normalizePosition(
                    getFromRecord<string>(teamPositions, member) || "",
                  );
                  if (mPos === "Unknown") return false;

                  const uplineResolved = getUplineInTeam(
                    member,
                    teamMembers,
                    teamUpLines,
                  );
                  return (
                    uplineResolved !== null &&
                    matchNames(uplineResolved, parentName)
                  );
                });
              };

              const renderRecursiveTeamNode = (
                parentName: string,
                depth = 1,
              ): React.ReactNode => {
                const directSubordinates = getDirectSubordinates(parentName);
                if (directSubordinates.length === 0) return null;

                return (
                  <div className="space-y-4 pl-3 border-l border-[#edecff]/70 ml-1.5 mt-2 animate-in fade-in duration-300">
                    {directSubordinates.map((sub, idx) => {
                      const subPos =
                        getFromRecord<string>(teamPositions, sub) ||
                        "Sales Agronomist";
                      const subProvince =
                        getFromRecord<string>(teamProvinces, sub) || "-";
                      const isCollapsed = collapsedNodes[sub] !== false;
                      const subSubs = getDirectSubordinates(sub);
                      const hasChildren = subSubs.length > 0;

                      return (
                        <div key={idx} className="space-y-4">
                          <div className="flex items-start">
                            <div
                              onClick={() => {
                                if (hasChildren) {
                                  setCollapsedNodes((prev) => ({
                                    ...prev,
                                    [sub]: prev[sub] === false,
                                  }));
                                }
                              }}
                              className={`bg-white hover:bg-slate-50 transition-all px-4 py-2.5 rounded-[18px] border border-slate-100 shadow-[0_8px_24px_rgba(21,75,226,0.06)] hover:shadow-[0_12px_28px_rgba(21,75,226,0.12)] flex-1 flex items-center justify-between ${hasChildren ? "cursor-pointer" : "cursor-default"}`}
                            >
                              <div className="flex-1 min-w-0 pr-2 text-left">
                                <p className="font-bold text-xs text-[#181a2c]">
                                  {normalizeName(sub)}
                                </p>
                                <div className="flex items-center gap-1.5 mt-1 flex-wrap">
                                  <span
                                    className={`text-[7.5px] font-extrabold uppercase tracking-wider px-1.5 py-0.5 rounded-full border ${
                                      depth === 1
                                        ? "text-emerald-600 bg-emerald-50 border-emerald-100"
                                        : depth === 2
                                          ? "text-blue-600 bg-blue-50 border-blue-100"
                                          : "text-purple-600 bg-purple-50 border-purple-100"
                                    }`}
                                  >
                                    {normalizePosition(subPos)}
                                  </span>
                                  {subProvince && subProvince !== "-" && (
                                    <span className="text-[7.5px] font-bold text-[#8E94B7] bg-slate-50 px-1.5 py-0.5 rounded-full border border-slate-100 uppercase tracking-wide">
                                      {subProvince}
                                    </span>
                                  )}
                                </div>
                              </div>

                              <div className="flex items-center gap-2 shrink-0 z-10 relative">
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    const emp = employees.find((emp) =>
                                      matchNames(emp.name, sub),
                                    );
                                    setEmployeeEditModal({
                                      isOpen: true,
                                      item: emp || {
                                        name: sub,
                                        position: subPos,
                                        province: subProvince,
                                      },
                                    });
                                  }}
                                  className="size-7 bg-[#edecff]/60 hover:bg-[#edecff] text-primary rounded-full border border-[#c4c5d8]/40 flex items-center justify-center transition-all shadow-none cursor-pointer"
                                  title="Edit Karyawan"
                                >
                                  <span className="material-symbols-outlined text-[13px]">
                                    edit
                                  </span>
                                </button>
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    const emp = employees.find((emp) =>
                                      matchNames(emp.name, sub),
                                    );
                                    setEmployeeDeleteModal({
                                      isOpen: true,
                                      item: emp || { name: sub },
                                    });
                                  }}
                                  className="size-7 bg-red-50 hover:bg-red-100 text-red-500 rounded-full border border-red-100 flex items-center justify-center transition-all shadow-none cursor-pointer"
                                  title="Hapus Karyawan"
                                >
                                  <span className="material-symbols-outlined text-[13px]">
                                    delete
                                  </span>
                                </button>
                                {hasChildren && (
                                  <span className="material-symbols-outlined text-[#8E94B7] text-md leading-none select-none ml-1">
                                    {isCollapsed
                                      ? "expand_more"
                                      : "expand_less"}
                                  </span>
                                )}
                              </div>
                            </div>
                          </div>

                          {hasChildren &&
                            !isCollapsed &&
                            renderRecursiveTeamNode(sub, depth + 1)}
                        </div>
                      );
                    })}
                  </div>
                );
              };

              const rootSubordinates = getDirectSubordinates(userData.name);

              return (
                <div className="animate-in fade-in slide-in-from-left-4 duration-300">
                  <div className="bg-white p-6 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.1)] mb-6 border border-[#edecff] animate-in fade-in slide-in-from-bottom-3 duration-300">
                    <div className="flex items-center justify-between gap-4 mb-5 pb-3 border-b border-[#edecff]">
                      <div>
                        <h3 className="text-xs font-bold text-[#181a2c] uppercase tracking-wider">
                          Hirarki Posisi Tim
                        </h3>
                      </div>
                      <button
                        disabled={userLevel === 1}
                        onClick={() =>
                          setEmployeeEditModal({
                            isOpen: true,
                            item: {
                              isAdd: true,
                              name: "",
                              position: "",
                              province: "",
                              email: "",
                              password: "",
                              upline: userData?.name || "",
                            },
                          })
                        }
                        className={`h-8 px-4 rounded-full font-bold text-[9px] uppercase tracking-wider flex items-center gap-1.5 transition-all shrink-0 ${
                          userLevel === 1
                            ? "bg-slate-200 text-slate-400 cursor-not-allowed"
                            : "bg-gradient-to-r from-teal-500 to-emerald-500 text-white hover:opacity-95 shadow-md active:scale-[97%] cursor-pointer"
                        }`}
                      >
                        <span className="material-symbols-outlined text-[13px]">
                          person_add
                        </span>
                        Tambah Anggota
                      </button>
                    </div>

                    {teamMembers.length > 0 ? (
                      <div className="space-y-4">
                        {/* Root Level 0 (Logged In User) */}
                        <div className="flex items-start gap-3">
                          <div className="size-7 rounded-full bg-primary/10 flex items-center justify-center font-bold text-[10px] text-primary shrink-0 shadow-sm border border-primary/10 mt-1.5">
                            0
                          </div>
                          <div
                            onClick={() => {
                              setCollapsedNodes((prev) => ({
                                ...prev,
                                [userData.name]: prev[userData.name] === false,
                              }));
                            }}
                            className={`bg-gradient-to-r from-[#edecff] to-sky-50/50 px-4 py-2.5 rounded-[18px] flex-1 relative overflow-hidden group shadow-[0_8px_24px_rgba(21,75,226,0.08)] hover:shadow-[0_12px_30px_rgba(21,75,226,0.14)] border-0 transition-all flex items-center justify-between ${rootSubordinates.length > 0 ? "cursor-pointer" : "cursor-default"}`}
                          >
                            <div className="absolute top-0 right-0 h-full w-24 bg-gradient-to-l from-primary/5 to-transparent pointer-events-none" />
                            <div className="flex-1 min-w-0">
                              <p className="font-bold text-xs text-[#181a2c]">
                                {normalizeName(userData.name)}
                              </p>
                              <div className="flex items-center gap-1.5 mt-1">
                                <span className="text-[7.5px] font-extrabold text-primary bg-primary/10 px-2 py-0.5 rounded-full border border-primary/10 uppercase tracking-wider">
                                  {normalizePosition(
                                    userData.position || "Business Analyst",
                                  )}
                                </span>
                                {userData.province &&
                                  userData.province !== "-" && (
                                    <span className="text-[7.5px] font-bold text-[#8E94B7] bg-slate-50 px-1.5 py-0.5 rounded-full border border-slate-100 uppercase tracking-wide">
                                      {userData.province}
                                    </span>
                                  )}
                              </div>
                            </div>

                            {/* Actions & Chevron */}
                            <div className="flex items-center gap-2 shrink-0 ml-3 z-10 relative">
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  const emp = employees.find((emp) =>
                                    matchNames(emp.name, userData.name),
                                  );
                                  setEmployeeEditModal({
                                    isOpen: true,
                                    item: emp || {
                                      name: userData.name,
                                      position: userData.position,
                                      province: userData.province || "",
                                      email: userData.email || "",
                                    },
                                  });
                                }}
                                className="size-7 bg-white/80 hover:bg-white text-primary rounded-full border border-[#edecff] flex items-center justify-center transition-all shadow-sm cursor-pointer"
                                title="Edit Karyawan"
                              >
                                <span className="material-symbols-outlined text-[14px]">
                                  edit
                                </span>
                              </button>
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  const emp = employees.find((emp) =>
                                    matchNames(emp.name, userData.name),
                                  );
                                  setEmployeeDeleteModal({
                                    isOpen: true,
                                    item: emp || { name: userData.name },
                                  });
                                }}
                                className="size-7 bg-red-50 hover:bg-red-100 text-red-500 rounded-full border border-red-100 flex items-center justify-center transition-all shadow-sm cursor-pointer"
                                title="Hapus Karyawan"
                                disabled={matchNames(
                                  userData.name,
                                  "Aditya Wiratama",
                                )}
                              >
                                <span className="material-symbols-outlined text-[14px]">
                                  delete
                                </span>
                              </button>
                              {rootSubordinates.length > 0 && (
                                <span className="material-symbols-outlined text-[#8E94B7] text-md leading-none select-none ml-1">
                                  {collapsedNodes[userData.name] !== false
                                    ? "expand_more"
                                    : "expand_less"}
                                </span>
                              )}
                            </div>
                          </div>
                        </div>

                        {/* Recursive Tree Render */}
                        {collapsedNodes[userData.name] === false &&
                          renderRecursiveTeamNode(userData.name, 1)}
                      </div>
                    ) : (
                      <p className="text-[9px] font-semibold text-[#8E94B7] text-center py-6 bg-slate-50/50 rounded-2xl border border-dashed border-[#edecff] uppercase tracking-wide">
                        Tidak ada anggota tim di bawah Anda. Silakan klik tombol
                        "Tambah Anggota" di atas untuk menambahkan.
                      </p>
                    )}
                  </div>
                </div>
              );
            })()}

          {partnerSubTab === "channel" && (
            <div className="animate-in fade-in slide-in-from-right-4 duration-300">
              {teamMembers.length > 1 && (
                <div className="mb-6">
                  <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wide ml-1 mb-2 block">
                    Pilih PIC / Tim
                  </label>
                  <div className="relative">
                    <select
                      value={mappingPic}
                      onChange={(e) => setMappingPic(e.target.value)}
                      className="w-full h-14 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-6 font-semibold text-xs text-[#111] outline-none focus:ring-1 focus:ring-primary/20 transition-all appearance-none cursor-pointer pr-12"
                    >
                      <option value="ALL_TEAM">Semua PIC</option>
                      {[...teamMembers]
                        .sort((a, b) =>
                          compareMembersByLevel(
                            a,
                            b,
                            teamLevels,
                            teamPositions,
                            userData,
                          ),
                        )
                        .map((picName, idx) => (
                          <option key={idx} value={picName}>
                            {picName}
                          </option>
                        ))}
                    </select>
                    <span className="material-symbols-outlined absolute right-6 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>
              )}

              <div className="mb-6 flex gap-2 overflow-x-auto pb-2 -mx-1 px-1 custom-scrollbar">
                {mappingCategories.map((cat) => (
                  <button
                    key={cat}
                    onClick={() => setMappingCategory(cat)}
                    className={`whitespace-nowrap px-4 py-2 rounded-full font-semibold text-[10.5px] uppercase tracking-wide transition-all ${mappingCategory === cat ? "bg-gradient-to-r from-primary to-cyan-400 text-white shadow-[0_4px_12px_rgba(21,75,226,0.2)]" : "bg-[#f4f2ff] text-[#8E94B7] hover:bg-[#edecff]"}`}
                  >
                    {cat}
                  </button>
                ))}
              </div>

              <div className="mb-4 flex items-center justify-between bg-white px-5 py-3.5 rounded-[20px] shadow-sm border border-[#edecff]">
                <div>
                  <h3 className="font-bold text-xs text-[#181a2c]">
                    Daftar Partner
                  </h3>
                </div>
                <button
                  onClick={() =>
                    setPartnerEditModal({
                      isOpen: true,
                      item: {
                        isAdd: true,
                        category: mappingCategory || "Kios",
                        name: "",
                        pic: "",
                      },
                    })
                  }
                  className="h-8 px-4 bg-gradient-to-r from-teal-500 to-emerald-500 text-white rounded-full font-bold text-[9px] uppercase tracking-wider flex items-center gap-1.5 hover:opacity-95 transition-all shadow-md active:scale-[97%] cursor-pointer"
                >
                  <span className="material-symbols-outlined text-[13px]">
                    add
                  </span>
                  Tambah Partner
                </button>
              </div>

              <div className="space-y-3">
                {displayedPartnerChannels.length === 0 ? (
                  <div className="py-12 bg-white rounded-[24px] shadow-sm flex flex-col justify-center items-center">
                    <span className="material-symbols-outlined text-[#8E94B7] text-3xl mb-2">
                      sentiment_dissatisfied
                    </span>
                    <p className="text-[11px] font-semibold text-[#8E94B7] uppercase tracking-wider">
                      Tidak Ada Data
                    </p>
                  </div>
                ) : (
                  displayedPartnerChannels.map((channel, i) => (
                    <div
                      key={i}
                      className="bg-white shadow-[0_4px_16px_rgba(21,75,226,0.06)] hover:shadow-[0_8px_24px_rgba(21,75,226,0.12)] p-4 rounded-[16px] flex items-center justify-between transition-all duration-250"
                    >
                      <div className="flex-1 pr-4">
                        <p className="font-semibold text-xs text-[#181a2c] leading-tight mb-0.5">
                          {channel.name}
                        </p>
                        <p className="text-[10px] font-bold text-[#8E94B7] flex items-center gap-1">
                          <span className="material-symbols-outlined text-[13px] text-[#8E94B7]/80">person</span>
                          {normalizeName(channel.pic) || "-"}
                        </p>
                        <div className="flex flex-wrap gap-1.5 mt-2">
                          {channel.category && (
                            <span className="text-[9px] font-extrabold bg-[#154be2]/10 text-primary px-2.5 py-0.5 rounded-full uppercase border border-[#154be2]/5">
                              {channel.category}
                            </span>
                          )}
                          {channel.group && (
                            <span className="text-[9px] font-extrabold bg-slate-100 text-slate-700 px-2.5 py-0.5 rounded-full uppercase">
                              {channel.group}
                            </span>
                          )}
                          {(channel.province || channel.area) && (
                            <span className="text-[9px] font-extrabold bg-emerald-50 text-emerald-700 px-2.5 py-0.5 rounded-full uppercase border border-emerald-100/30">
                              {channel.province || channel.area}
                            </span>
                          )}
                        </div>
                      </div>
                      <div className="flex gap-2">
                        <button
                          onClick={() =>
                            setPartnerEditModal({ isOpen: true, item: channel })
                          }
                          className="size-9 bg-[#edecff] text-primary rounded-full border border-[#c4c5d8] flex items-center justify-center hover:bg-[#e6e6ff] transition-all shadow-none cursor-pointer"
                        >
                          <span className="material-symbols-outlined text-[16px]">
                            edit
                          </span>
                        </button>
                        <button
                          onClick={() =>
                            setPartnerDeleteModal({
                              isOpen: true,
                              item: channel,
                            })
                          }
                          className="size-9 bg-red-50 text-red-500 rounded-full border border-red-100 flex items-center justify-center hover:bg-red-100 transition-all shadow-none cursor-pointer"
                          title="Hapus Partner"
                        >
                          <span className="material-symbols-outlined text-[16px]">
                            delete
                          </span>
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          )}

          <PartnerEditModal
            isOpen={partnerEditModal.isOpen}
            item={partnerEditModal.item}
            onClose={() => setPartnerEditModal({ isOpen: false, item: null })}
            onSave={handleEditPartnerSave}
            isSaving={isActionLoading}
            activeEmployees={myActiveSubordinates}
            allProvinces={availableProvinces}
            allCategories={allCategories}
            userData={userData}
            allGroups={availableGroups}
          />
          <PartnerDeleteModal
            isOpen={partnerDeleteModal.isOpen}
            onClose={() => setPartnerDeleteModal({ isOpen: false, item: null })}
            onConfirm={handleDeletePartnerConfirm}
            isProcessing={isActionLoading}
            itemName={partnerDeleteModal.item?.name}
          />

          <EmployeeEditModal
            isOpen={employeeEditModal.isOpen}
            item={employeeEditModal.item}
            onClose={() => setEmployeeEditModal({ isOpen: false, item: null })}
            onSave={handleEditEmployeeSave}
            isSaving={isActionLoading}
            allEmployeeNames={employees.map((emp) => emp.name)}
            userData={userData}
            allProvinces={availableProvinces}
            accessRules={accessRules}
          />
          <EmployeeDeleteModal
            isOpen={employeeDeleteModal.isOpen}
            onClose={() =>
              setEmployeeDeleteModal({ isOpen: false, item: null })
            }
            onConfirm={handleDeleteEmployeeConfirm}
            isProcessing={isActionLoading}
            itemName={employeeDeleteModal.item?.name}
          />


        </div>
      )}

      {activeTab === "summary" && (
        <div className="animate-in fade-in slide-in-from-right-4 duration-300">
          <div className="mb-4 ml-1 flex items-center gap-3">
            <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
              Stock <span className="text-primary font-bold">Summary</span>
            </h1>
            {filteredSummaryData.length > 0 && (
              <button
                onClick={handleDownloadSummaryExcel}
                className="h-8.5 w-16 bg-gradient-to-r from-emerald-500 to-teal-600 text-white rounded-full flex items-center justify-center hover:opacity-95 transition-all shadow-md active:scale-[97%] cursor-pointer shrink-0"
                title="Download Excel"
              >
                <span className="material-symbols-outlined text-[18px]">
                  download
                </span>
              </button>
            )}
          </div>

          <div className="flex flex-col gap-3.5 mb-6">
            <div className="flex flex-col md:flex-row gap-4 md:gap-6 md:items-stretch">
              {filteredSummaryData.length > 0 && renderPogKpiCard(true)}

              <div
                className={`w-full md:w-64 lg:w-72 xl:w-80 shrink-0 transition-all ${isSummaryFilterOpen ? "block" : "hidden md:block"}`}
              >
                <div className="bg-white p-4 md:px-6 md:py-5 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.22)] animate-in fade-in slide-in-from-top-3 duration-300 h-full flex flex-col justify-center">
                  <div className="flex flex-row md:flex-col gap-2 md:gap-3.5">
                    <div className="flex-1 min-w-0">
                      <label className="text-[9px] text-[#8E94B7] font-bold uppercase tracking-wider block mb-1.5 truncate ml-1">
                        Group By
                      </label>
                      <div className="relative">
                        <select
                          value={summaryGroupBy}
                          onChange={(e) => setSummaryGroupBy(e.target.value)}
                          className="w-full h-10 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-4 font-semibold text-[10.5px] text-primary outline-none truncate appearance-none"
                        >
                          <option value="hybrid">Hybrid</option>
                          <option value="subordinate">Team</option>
                          <option value="area">Area</option>
                          <option value="category">Category</option>
                          <option value="crops">Crops</option>
                        </select>
                        <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-primary text-base pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>
                    <div className="flex-1 min-w-0">
                      <label className="text-[9px] text-[#8E94B7] font-bold uppercase tracking-wider block mb-1.5 truncate ml-1">
                        Sub
                      </label>
                      <div className="relative">
                        <select
                          value={summarySubGroupBy}
                          onChange={(e) => setSummarySubGroupBy(e.target.value)}
                          className="w-full h-10 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-4 font-semibold text-[10.5px] text-primary outline-none truncate appearance-none"
                        >
                          <option value="channel">Channel</option>
                          <option value="hybrid">Hybrid</option>
                          <option value="subordinate">Team</option>
                          <option value="area">Area</option>
                          <option value="category">Category</option>
                          <option value="crops">Crops</option>
                        </select>
                        <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-primary text-base pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>
                    <div className="flex-1 min-w-0">
                      <label className="text-[9px] text-[#8E94B7] font-bold uppercase tracking-wider block mb-1.5 truncate ml-1">
                        Crop
                      </label>
                      <div className="relative">
                        <select
                          value={filterBelowCrop}
                          onChange={(e) => setFilterBelowCrop(e.target.value)}
                          className="w-full h-10 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-4 font-semibold text-[10.5px] text-primary outline-none truncate appearance-none"
                        >
                          {availableCrops.map((crop) => (
                            <option key={crop} value={crop}>
                              {crop}
                            </option>
                          ))}
                        </select>
                        <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-primary text-base pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {filteredSummaryData.length > 0 && (
              <div className="-mx-5 md:mx-0 md:bg-transparent md:border-none md:shadow-none md:divide-y-0 bg-white overflow-hidden md:overflow-visible rounded-[48px] md:rounded-none pt-4 md:pt-0 pb-4 md:pb-0 mb-8 shadow-[0_4px_44px_rgba(24,26,44,0.15)] border border-[#edecff] divide-y divide-[#edecff] transition-all md:grid md:grid-cols-2 md:gap-4 xl:grid-cols-2">
                {summaryGroupBy === "subordinate"
                  ? filteredSummaryData.map((row) =>
                      renderRecursiveSummaryRow(row),
                    )
                  : filteredSummaryData.map((row, i) => (
                      <div
                        key={i}
                        className="p-0 overflow-hidden transition-all md:bg-white md:rounded-[32px] md:shadow-[0_8px_32px_rgba(21,75,226,0.18)] md:border md:border-[#edecff]"
                      >
                        <div
                          className={`flex justify-between items-center ${row.isExpandable ? "cursor-pointer" : ""} px-5 py-4 pb-2 hover:bg-slate-50/60 transition-colors`}
                          onClick={() =>
                            row.isExpandable && toggleRow(row.name)
                          }
                        >
                          <div className="flex flex-col">
                            <span className="font-semibold text-xs md:text-sm text-[#181a2c] uppercase flex items-center gap-1.5">
                              {row.isExpandable && (
                                <span className="material-symbols-outlined text-primary text-[20px]">
                                  {expandedRows[row.name]
                                    ? "keyboard_arrow_down"
                                    : "keyboard_arrow_right"}
                                </span>
                              )}
                              {row.name}
                            </span>
                          </div>
                          <div className="flex flex-col items-end">
                            <span className="font-bold text-sm text-primary">
                              {formatOverviewVal(row.selectedTotal, overviewUseMt).valueStr}
                            </span>
                            <span className="text-[8px] text-[#8E94B7] uppercase tracking-widest font-bold">
                              Total {formatOverviewVal(row.selectedTotal, overviewUseMt).unit}
                            </span>
                          </div>
                        </div>

                        <div className="mx-5 mb-3 mt-2 flex divide-x divide-white/20 bg-primary shadow-[0_12px_32px_rgba(21,75,226,0.35)] rounded-[14px] overflow-hidden">
                          {selectedClusters.map((clusterKey) => {
                            if (
                              clusterKey === "Uncategorized" &&
                              (!row[clusterKey] || row[clusterKey] === 0)
                            )
                              return null;
                            const clusterConfig = CLUSTER_CONFIG.find(
                              (c) => c.key === clusterKey,
                            );
                            return (
                              <div
                                key={clusterKey}
                                className="flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center hover:bg-white/10 transition-colors"
                              >
                                <span className="text-[8px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                                  {clusterConfig?.label || clusterKey}
                                </span>
                                <span className="font-semibold text-[10.5px] truncate w-full text-white">
                                  {formatOverviewVal(row[clusterKey], overviewUseMt).valueStr}
                                </span>
                              </div>
                            );
                          })}
                        </div>

                        {expandedRows[row.name] && row.children?.length > 0 && (
                          <div className="pb-4 pt-3 border-t border-[#edecff] flex flex-col gap-3.5 px-5 bg-slate-50/50">
                            {summarySubGroupBy === "subordinate"
                              ? row.children.map((child: any) =>
                                  renderRecursiveSubordinate(child),
                                )
                              : row.children.map((child, j) => {
                                  const isChildZeroTeam =
                                    summarySubGroupBy === "subordinate" &&
                                    child.selectedTotal === 0;
                                  return (
                                    <div
                                      key={`${i}-${j}`}
                                      className={`flex flex-col p-3.5 rounded-[18px] transition-all duration-200 ${
                                        isChildZeroTeam
                                          ? "bg-red-50/70 border border-red-200/60 shadow-[0_10px_28px_rgba(239,68,68,0.12)]"
                                          : "bg-[#fbfaff] shadow-[0_10px_28px_rgba(21,75,226,0.18)]"
                                      }`}
                                    >
                                      <div className="flex justify-between items-center mb-2 px-1">
                                        <div className="flex flex-col">
                                          <span className="font-bold text-[11px] text-[#181a2c] uppercase pr-2 flex items-center gap-1.5 flex-wrap">
                                            <span>
                                              {renderMaybeChannelName(
                                                child.name,
                                              )}
                                            </span>
                                            {isChildZeroTeam && (
                                              <span className="text-[7.5px] bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full font-bold uppercase tracking-wider">
                                                No Activity
                                              </span>
                                            )}
                                            {summarySubGroupBy === "channel" &&
                                              child.category && (
                                                <span className="text-[8px] bg-primary/10 text-primary px-1.5 py-0.5 rounded-full font-bold">
                                                  {child.category}
                                                </span>
                                              )}
                                          </span>
                                        </div>
                                        <span
                                          className={`font-bold text-[11.5px] shrink-0 ${isChildZeroTeam ? "text-red-600" : "text-[#181a2c]"}`}
                                        >
                                          {formatOverviewVal(child.selectedTotal, overviewUseMt).valueStr}{" "}
                                          <span className="text-[8.5px] text-[#8E94B7]">
                                            {formatOverviewVal(child.selectedTotal, overviewUseMt).unit}
                                          </span>
                                        </span>
                                      </div>
                                      <div
                                        className={`flex w-full divide-x rounded-[14px] overflow-hidden ${
                                          isChildZeroTeam
                                            ? "divide-red-200 bg-red-100/40"
                                            : "divide-primary/10 bg-primary/8"
                                        }`}
                                      >
                                        {selectedClusters.map((clusterKey) => {
                                          if (
                                            clusterKey === "Uncategorized" &&
                                            (!child[clusterKey] ||
                                              child[clusterKey] === 0)
                                          )
                                            return null;
                                          const clusterConfig =
                                            CLUSTER_CONFIG.find(
                                              (c) => c.key === clusterKey,
                                            );
                                          return (
                                            <div
                                              key={clusterKey}
                                              className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                                                isChildZeroTeam
                                                  ? "hover:bg-red-200/30"
                                                  : "hover:bg-primary/5"
                                              }`}
                                            >
                                              <span
                                                className={`text-[7.5px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-700/60"
                                                    : "text-primary/70"
                                                }`}
                                              >
                                                {clusterConfig?.label ||
                                                  clusterKey}
                                              </span>
                                              <span
                                                className={`font-black text-[10px] truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-700"
                                                    : "text-primary"
                                                }`}
                                              >
                                                {formatOverviewVal(child[clusterKey], overviewUseMt).valueStr}
                                              </span>
                                            </div>
                                          );
                                        })}
                                      </div>
                                    </div>
                                  );
                                })}
                          </div>
                        )}
                      </div>
                    ))}
              </div>
            )}

            {filteredSummaryData.length === 0 && (
              <div className="py-16 text-center flex flex-col items-center bg-white rounded-[24px] border border-[#edecff] shadow-sm">
                <span className="material-symbols-outlined text-[40px] text-[#8E94B7] mb-4">
                  inventory_2
                </span>
                <p className="font-semibold text-[#8E94B7] text-xs">
                  No data available for the selected filters.
                </p>
              </div>
            )}
          </div>
        </div>
      )}

      {activeTab === "pog" && (
        <div className="animate-in fade-in slide-in-from-right-4 duration-300">
          <div className="mb-4 ml-1">
            <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
              POG <span className="text-primary font-bold">Analytics</span>
            </h1>
          </div>

          <div className="flex flex-col gap-3.5 mb-6">
            <div className="flex flex-col md:flex-row gap-4 md:gap-6 md:items-stretch">
              {renderPogKpiCard(false)}

              <div
                className={`w-full md:w-64 lg:w-72 xl:w-80 shrink-0 transition-all ${isPogFilterOpen ? "block" : "hidden md:block"}`}
              >
                <div className="bg-white p-4 md:px-6 md:py-5 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.22)] animate-in fade-in slide-in-from-top-3 duration-300 h-full flex flex-col justify-center">
                  <div className="flex flex-row md:flex-col gap-2 md:gap-3.5">
                    <div className="flex-1 min-w-0">
                      <label className="text-[9px] text-[#8E94B7] font-bold uppercase tracking-wider block mb-1.5 truncate ml-1">
                        Group By
                      </label>
                      <div className="relative">
                        <select
                          value={pogGroupBy}
                          onChange={(e) => setPogGroupBy(e.target.value)}
                          className="w-full h-10 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-4 font-semibold text-[10.5px] text-primary outline-none truncate appearance-none pr-7"
                        >
                          <option value="hybrid">Hybrid</option>
                          <option value="subordinate">Team</option>
                          <option value="area">Area</option>
                          <option value="category">Category</option>
                          <option value="crops">Crops</option>
                        </select>
                        <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-primary text-base pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>
                    <div className="flex-1 min-w-0">
                      <label className="text-[9px] text-[#8E94B7] font-bold uppercase tracking-wider block mb-1.5 truncate ml-1">
                        Sub
                      </label>
                      <div className="relative">
                        <select
                          value={pogSubGroupBy}
                          onChange={(e) => setPogSubGroupBy(e.target.value)}
                          className="w-full h-10 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-4 font-semibold text-[10.5px] text-primary outline-none truncate appearance-none pr-7"
                        >
                          <option value="channel">Channel</option>
                          <option value="hybrid">Hybrid</option>
                          <option value="subordinate">Team</option>
                          <option value="area">Area</option>
                          <option value="category">Category</option>
                          <option value="crops">Crops</option>
                        </select>
                        <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-primary text-base pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>
                    <div className="flex-1 min-w-0">
                      <label className="text-[9px] text-[#8E94B7] font-bold uppercase tracking-wider block mb-1.5 truncate ml-1">
                        Crop
                      </label>
                      <div className="relative">
                        <select
                          value={filterBelowCrop}
                          onChange={(e) => setFilterBelowCrop(e.target.value)}
                          className="w-full h-10 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-4 font-semibold text-[10.5px] text-primary outline-none truncate appearance-none pr-7"
                        >
                          {availableCrops.map((crop) => (
                            <option key={crop} value={crop}>
                              {crop}
                            </option>
                          ))}
                        </select>
                        <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-primary text-base pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {aggregatedPogData.length > 0 && (
              <div className="-mx-5 md:mx-0 md:bg-transparent md:border-none md:shadow-none md:divide-y-0 bg-white overflow-hidden md:overflow-visible rounded-[48px] md:rounded-none pt-4 md:pt-0 pb-4 md:pb-0 mb-8 shadow-[0_4px_44px_rgba(24,26,44,0.15)] border border-[#edecff] divide-y divide-[#edecff] transition-all md:grid md:grid-cols-2 md:gap-4 xl:grid-cols-2">
                {pogGroupBy === "subordinate"
                  ? aggregatedPogData.map((row) => renderRecursivePogRow(row))
                  : aggregatedPogData.map((row, i) => (
                      <div
                        key={i}
                        className="p-0 overflow-hidden transition-all md:bg-white md:rounded-[32px] md:shadow-[0_8px_32px_rgba(21,75,226,0.18)] md:border md:border-[#edecff]"
                      >
                        <div
                          className={`flex justify-between items-center ${row.isExpandable ? "cursor-pointer" : ""} px-5 py-4 pb-2 hover:bg-slate-50/60 transition-colors`}
                          onClick={() =>
                            row.isExpandable && togglePogRow(row.name)
                          }
                        >
                          <div className="flex flex-col">
                            <span className="font-semibold text-xs md:text-sm text-[#181a2c] uppercase flex items-center gap-1.5">
                              {row.isExpandable && (
                                <span className="material-symbols-outlined text-primary text-[20px]">
                                  {pogExpandedRows[row.name]
                                    ? "keyboard_arrow_down"
                                    : "keyboard_arrow_right"}
                                </span>
                              )}
                              {row.name}
                            </span>
                          </div>
                          <div className="flex flex-col items-end">
                            <span className="font-bold text-sm text-primary">
                              {formatOverviewVal(row.pog, overviewUseMt).valueStr}
                            </span>
                            <span className="text-[8px] text-[#8E94B7] uppercase tracking-widest font-bold">
                              POG ({formatOverviewVal(row.pog, overviewUseMt).unit})
                            </span>
                          </div>
                        </div>

                        <div className="mx-5 mb-3 mt-2 flex flex-row gap-1.5 md:gap-2">
                          {/* Table 1: Opening Inv & End of Inv */}
                          <div className="flex-[2] flex divide-x divide-white/20 bg-primary/95 shadow-[0_12px_32px_rgba(21,75,226,0.25)] rounded-[14px] overflow-hidden">
                            <div className="flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center hover:bg-white/10 transition-colors">
                              <span className="text-[8px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                                Opening Inv
                              </span>
                              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                                {formatOverviewVal(row.lastQty, overviewUseMt).valueStr}
                              </span>
                            </div>
                            <div className="flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center hover:bg-white/10 transition-colors">
                              <span className="text-[8px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                                End of Inv
                              </span>
                              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                                {formatOverviewVal(row.currentQty, overviewUseMt).valueStr}
                              </span>
                            </div>
                          </div>

                          {/* Table 2: Stock in, idle stock, POG */}
                          <div className="flex-[3] flex divide-x divide-white/20 bg-primary shadow-[0_12px_32px_rgba(21,75,226,0.35)] rounded-[14px] overflow-hidden">
                            <div className="flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center hover:bg-white/10 transition-colors">
                              <span className="text-[8px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                                Stock in
                              </span>
                              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                                {formatOverviewVal(row.sellIn, overviewUseMt).valueStr}
                              </span>
                            </div>
                            <div className="flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center hover:bg-white/10 transition-colors">
                              <span className="text-[8px] font-bold uppercase tracking-wider text-amber-200 mb-0.5 truncate w-full">
                                idle stock
                              </span>
                              <span className="font-semibold text-[10.5px] truncate w-full text-amber-100">
                                {formatOverviewVal(row.idleStock, overviewUseMt).valueStr}
                              </span>
                            </div>
                            <div className="flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center hover:bg-white/10 transition-colors">
                              <span className="text-[8px] font-bold uppercase tracking-wider text-cyan-200 mb-0.5 truncate w-full">
                                POG
                              </span>
                              <span className="font-semibold text-[10.5px] truncate w-full text-cyan-100 font-extrabold">
                                {formatOverviewVal(row.pog, overviewUseMt).valueStr}
                              </span>
                            </div>
                          </div>
                        </div>

                        {pogExpandedRows[row.name] &&
                          row.children?.length > 0 && (
                            <div className="pb-4 pt-3 border-t border-[#edecff] flex flex-col gap-3.5 px-5 bg-slate-50/50">
                              {pogSubGroupBy === "subordinate"
                                ? row.children.map((child: any) =>
                                    renderRecursivePogSubordinate(child),
                                  )
                                : row.children.map((child, j) => {
                                    const isChildZeroTeam =
                                      child.lastQty === 0 &&
                                      child.sellIn === 0 &&
                                      child.sellOut === 0 &&
                                      child.totalInv === 0 &&
                                      child.currentQty === 0 &&
                                      (child.idleStock || 0) === 0;
                                    return (
                                      <div
                                        key={`${i}-${j}`}
                                        className={`flex flex-col p-3.5 rounded-[18px] transition-all duration-200 ${
                                          isChildZeroTeam
                                            ? "bg-red-50/70 border border-red-200/60 shadow-[0_10px_28px_rgba(239,68,68,0.12)]"
                                            : "bg-[#fbfaff] shadow-[0_10px_28px_rgba(21,75,226,0.18)]"
                                        }`}
                                      >
                                        <div className="flex justify-between items-center mb-2 px-1 flex-wrap gap-2">
                                          <div className="flex flex-col">
                                            <span className="font-bold text-[11px] text-[#181a2c] uppercase pr-2 flex items-center gap-1.5 flex-wrap">
                                              <span>
                                                {renderMaybeChannelName(
                                                  child.name,
                                                )}
                                              </span>
                                              {isChildZeroTeam && (
                                                <span className="text-[7.5px] bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full font-bold uppercase tracking-wider">
                                                  No Activity
                                                </span>
                                              )}
                                              {pogSubGroupBy === "channel" &&
                                                child.category && (
                                                  <span className="text-[8px] bg-primary/10 text-primary px-1.5 py-0.5 rounded-full font-bold">
                                                    {child.category}
                                                  </span>
                                                )}
                                            </span>
                                          </div>
                                          <span
                                            className={`font-bold text-[11.5px] shrink-0 ${isChildZeroTeam ? "text-red-600" : "text-primary"}`}
                                          >
                                            {formatOverviewVal(child.pog, overviewUseMt).valueStr}{" "}
                                            <span className="text-[8.5px] text-[#8E94B7]">
                                              POG ({formatOverviewVal(child.pog, overviewUseMt).unit})
                                            </span>
                                          </span>
                                        </div>
                                        <div className="flex flex-row w-full gap-1.5 md:gap-2">
                                          {/* Table 1: Opening Inv & End of Inv */}
                                          <div
                                            className={`flex-[2] flex divide-x rounded-[14px] overflow-hidden ${
                                              isChildZeroTeam
                                                ? "divide-red-200 bg-red-100/40"
                                                : "divide-primary/10 bg-primary/5"
                                            }`}
                                          >
                                            <div
                                              className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                                                isChildZeroTeam
                                                  ? "hover:bg-red-200/30"
                                                  : "hover:bg-primary/5"
                                              }`}
                                            >
                                              <span
                                                className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                                                  isChildZeroTeam
                                                    ? "text-red-700/60"
                                                    : "text-[#8E94B7]"
                                                }`}
                                              >
                                                Opening Inv
                                              </span>
                                              <span
                                                className={`font-black text-[10px] truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-700"
                                                    : "text-[#181a2c]"
                                                }`}
                                              >
                                                {formatOverviewVal(child.lastQty, overviewUseMt).valueStr}
                                              </span>
                                            </div>
                                            <div
                                              className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                                                isChildZeroTeam
                                                  ? "hover:bg-red-200/30"
                                                  : "hover:bg-primary/5"
                                              }`}
                                            >
                                              <span
                                                className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                                                  isChildZeroTeam
                                                    ? "text-red-700/60"
                                                    : "text-[#1d4ed8]/75"
                                                }`}
                                              >
                                                End of Inv
                                              </span>
                                              <span
                                                className={`font-black text-[10px] truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-700"
                                                    : "text-[#1d4ed8]"
                                                }`}
                                              >
                                                {formatOverviewVal(child.currentQty, overviewUseMt).valueStr}
                                              </span>
                                            </div>
                                          </div>

                                          {/* Table 2: Stock in, idle stock, POG */}
                                          <div
                                            className={`flex-[3] flex divide-x rounded-[14px] overflow-hidden ${
                                              isChildZeroTeam
                                                ? "divide-red-200 bg-red-200/45"
                                                : "divide-primary/10 bg-primary/10 border border-primary/10"
                                            }`}
                                          >
                                            <div
                                              className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                                                isChildZeroTeam
                                                  ? "hover:bg-red-300/30"
                                                  : "hover:bg-primary/15"
                                              }`}
                                            >
                                              <span
                                                className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                                                  isChildZeroTeam
                                                    ? "text-red-800/70"
                                                    : "text-[#154be2]/80"
                                                }`}
                                              >
                                                Stock in
                                              </span>
                                              <span
                                                className={`font-black text-[10px] truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-800"
                                                    : "text-[#154be2]"
                                                }`}
                                              >
                                                {formatOverviewVal(child.sellIn, overviewUseMt).valueStr}
                                              </span>
                                            </div>
                                            <div
                                              className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                                                isChildZeroTeam
                                                  ? "hover:bg-red-300/30"
                                                  : "hover:bg-amber-100/60"
                                              }`}
                                            >
                                              <span
                                                className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                                                  isChildZeroTeam
                                                    ? "text-red-800/70"
                                                    : "text-amber-800"
                                                }`}
                                              >
                                                idle stock
                                              </span>
                                              <span
                                                className={`font-black text-[10px] truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-800"
                                                    : "text-amber-700"
                                                }`}
                                              >
                                                {formatOverviewVal(child.idleStock, overviewUseMt).valueStr}
                                              </span>
                                            </div>
                                            <div
                                              className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                                                isChildZeroTeam
                                                  ? "hover:bg-red-300/30"
                                                  : "hover:bg-emerald-100/60"
                                              }`}
                                            >
                                              <span
                                                className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                                                  isChildZeroTeam
                                                    ? "text-red-800/70"
                                                    : "text-emerald-800"
                                                }`}
                                              >
                                                POG
                                              </span>
                                              <span
                                                className={`font-black text-[10px] truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-800"
                                                    : "text-emerald-700 font-extrabold"
                                                }`}
                                              >
                                                {formatOverviewVal(child.pog, overviewUseMt).valueStr}
                                              </span>
                                            </div>
                                          </div>
                                        </div>
                                      </div>
                                    );
                                  })}
                            </div>
                          )}
                      </div>
                    ))}
              </div>
            )}

            {aggregatedPogData.length === 0 && (
              <div className="py-16 text-center flex flex-col items-center bg-white rounded-[24px] border border-[#edecff] shadow-sm">
                <span className="material-symbols-outlined text-[40px] text-[#8E94B7] mb-4">
                  analytics
                </span>
                <p className="font-semibold text-[#8E94B7] text-xs">
                  No data available for the selected filters.
                </p>
              </div>
            )}
          </div>
        </div>
      )}

      {activeTab === "access" && (
        <div className="animate-in fade-in slide-in-from-right-4 duration-300">
          <div className="mb-6 ml-1 flex flex-col gap-1">
            <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
              Access Control Menu
            </h1>
            <p className="text-[#8E94B7] text-[11px] font-semibold tracking-wide">
              Level permissions and access rights mapping
            </p>
          </div>
          
          <div className="bg-white rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.12)] border border-[#154be2]/5 overflow-hidden">
            <div className="p-6 border-b border-[#f1f5f9]">
              <div className="flex items-center gap-2">
                <div className="size-8 rounded-lg bg-primary/10 flex items-center justify-center">
                  <span className="material-symbols-outlined text-primary text-[20px]">
                    admin_panel_settings
                  </span>
                </div>
                <div>
                  <h3 className="text-sm font-bold text-[#181a2c]">System Access Rights</h3>
                  <p className="text-[10px] font-semibold text-[#8E94B7] mt-0.5">Matrix of features available for each organizational level</p>
                </div>
              </div>
            </div>
            
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse min-w-[700px]">
                <thead>
                  <tr className="bg-[#fbfaff]">
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Position</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Data Partner</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Stock Summary</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">POG Tracking</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Overview Tab</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Temp Tab</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Access Menu</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9] text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#f1f5f9]">
                  {allPositionsList.map((position) => (
                    <tr key={position} className="hover:bg-slate-50/50 transition-colors">
                      <td className="px-5 py-4">
                        <span className="text-[11px] font-semibold text-[#181a2c] bg-slate-100 px-2 py-1 rounded-md">{position}</span>
                      </td>
                      <td className="px-5 py-4">{renderAccessCheckbox(position, 'partner')}</td>
                      <td className="px-5 py-4">{renderAccessCheckbox(position, 'stock')}</td>
                      <td className="px-5 py-4">{renderAccessCheckbox(position, 'pog')}</td>
                      <td className="px-5 py-4">{renderAccessCheckbox(position, 'overview')}</td>
                      <td className="px-5 py-4">{renderAccessCheckbox(position, 'temp')}</td>
                      <td className="px-5 py-4">{renderAccessCheckbox(position, 'access')}</td>
                      <td className="px-5 py-4 text-right">
                        <button onClick={() => removeAccessRule(position)} className="text-red-500 hover:bg-red-50 p-1.5 rounded-lg transition-colors" title="Hapus Rule">
                          <span className="material-symbols-outlined text-[16px]">delete</span>
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            
            <div className="p-5 border-t border-[#f1f5f9] bg-[#fbfaff] flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div className="flex items-start gap-3 flex-1">
                <span className="material-symbols-outlined text-[#8E94B7] text-[18px]">info</span>
                <p className="text-[10px] font-semibold text-[#8E94B7] leading-relaxed">
                  <span className="text-[#181a2c] font-bold uppercase tracking-wider block mb-1">General Access Rules</span>
                  All levels have access to the <strong className="text-primary">Home</strong>, <strong className="text-primary">Data Partner</strong>, <strong className="text-primary">Stock Summary</strong>, and <strong className="text-primary">POG Tracking</strong> tabs. The data visible within these tabs is automatically filtered based on the user's <strong className="text-primary">Data Visibility</strong> level. The <strong className="text-primary">Access Menu</strong> tab is strictly limited to authorized system administrators.
                </p>
              </div>
              
              <div className="flex items-center justify-end gap-3 w-full md:w-auto mt-2 md:mt-0">
                {accessSaveSuccess && (
                  <div className="flex items-center gap-1.5 text-emerald-600 bg-emerald-50 px-3 py-1.5 rounded-lg text-xs font-bold animate-in fade-in slide-in-from-bottom-2">
                    <span className="material-symbols-outlined text-[16px]">check_circle</span>
                    Saved
                  </div>
                )}
                <button
                  onClick={handleSaveAccessRules}
                  disabled={isSavingAccess}
                  className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-[#154be2] hover:bg-[#154be2]/90 text-white px-6 py-2.5 rounded-xl text-xs font-bold transition-all shadow-sm disabled:opacity-70 disabled:cursor-not-allowed cursor-pointer"
                >
                  {isSavingAccess ? (
                    <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                  ) : (
                    <span className="material-symbols-outlined text-[18px]">save</span>
                  )}
                  {isSavingAccess ? "Saving..." : "Save Role Access"}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {activeTab === "temp" && (
        <div className="animate-in fade-in slide-in-from-right-4 duration-300 animate-in">
          {consolidationSuccessMsg && (
            <div className="mb-4 bg-emerald-50 border border-emerald-200 text-emerald-800 rounded-2xl px-5 py-3 text-xs font-bold flex items-center gap-2 shadow-sm animate-in fade-in slide-in-from-top-2 duration-300">
              <span className="material-symbols-outlined text-emerald-600">
                check_circle
              </span>
              {consolidationSuccessMsg}
            </div>
          )}

          <div className="mb-6 flex flex-col lg:flex-row lg:items-center justify-between gap-4 ml-1">
            <div>
              <h1 className="text-xl font-bold text-[#181a2c] tracking-tight">
                Review{" "}
                <span className="text-primary font-bold">Data (Temporary)</span>
              </h1>
              <p className="text-xs text-[#8E94B7] mt-1 font-semibold">
                Menampilkan data checker, channel, hybrid, dan lot no secara
                komprehensif.
              </p>
            </div>

            <div className="flex flex-col sm:flex-row items-center gap-3 w-full lg:w-auto shrink-0">
              {/* Preview Toggle Button */}
              <button
                type="button"
                id="btn-temp-proceed-consolidate"
                onClick={() => {
                  setIsTempProceeded(!isTempProceeded);
                  setExpandedTempRowId(null); // Close any expanded row during toggle
                }}
                className={`w-full sm:w-auto h-10 px-5 rounded-full font-extrabold text-xs flex items-center justify-center gap-2 transition-all cursor-pointer active:scale-[0.98] ${
                  isTempProceeded
                    ? "bg-amber-500 text-white shadow-[0_4px_14px_rgba(245,158,11,0.3)] hover:bg-amber-600"
                    : "bg-[#181a2c] text-white shadow-[0_4px_14px_rgba(24,26,44,0.15)] hover:bg-[#252841]"
                }`}
              >
                <span className="material-symbols-outlined text-[18px]">
                  {isTempProceeded ? "visibility_off" : "visibility"}
                </span>
                {isTempProceeded ? "Matikan Preview" : "Preview Konsolidasi"}
              </button>

              {/* Real Database Process Button */}
              <button
                type="button"
                id="btn-temp-process-db-consolidate"
                disabled={isConsolidatingDb}
                onClick={handleConsolidateDatabase}
                className={`w-full sm:w-auto h-10 px-5 rounded-full font-extrabold text-xs flex items-center justify-center gap-2 transition-all cursor-pointer active:scale-[0.98] bg-emerald-600 text-white shadow-[0_4px_14px_rgba(16,185,129,0.3)] hover:bg-emerald-700 disabled:opacity-50 disabled:cursor-not-allowed`}
              >
                <span
                  className={`material-symbols-outlined text-[18px] ${isConsolidatingDb ? "animate-spin" : ""}`}
                >
                  {isConsolidatingDb ? "sync" : "auto_mode"}
                </span>
                {isConsolidatingDb ? "Memproses..." : "Proses Konsolidasi"}
              </button>

              {/* Quick Search Container */}
              <div className="relative w-full sm:w-64 shrink-0">
                <span className="material-symbols-outlined absolute left-4 top-1/2 -translate-y-1/2 text-gray-400 text-[18px]">
                  search
                </span>
                <input
                  type="text"
                  placeholder="Cari data..."
                  value={tempSearchQuery}
                  onChange={(e) => setTempSearchQuery(e.target.value)}
                  className="w-full h-10 bg-white border border-[#edecff] shadow-sm rounded-full pl-11 pr-10 font-bold text-xs text-[#181a2c] outline-none focus:ring-1 focus:ring-primary/20 transition-all placeholder:text-gray-400"
                />
                {tempSearchQuery && (
                  <button
                    type="button"
                    onClick={() => setTempSearchQuery("")}
                    className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 p-1 flex items-center justify-center cursor-pointer transition-colors"
                  >
                    <span className="material-symbols-outlined text-[16px]">
                      close
                    </span>
                  </button>
                )}
              </div>
            </div>
          </div>

          {/* Petunjuk Konsolidasi & Sinkronisasi */}
          <div className="mb-6 bg-amber-50/70 border border-amber-200/80 rounded-2xl p-5 shadow-sm ml-1 text-[#181a2c]">
            <div className="flex gap-3">
              <span className="material-symbols-outlined text-[22px] text-amber-600 mt-0.5">
                info
              </span>
              <div className="flex-1">
                <h3 className="text-xs font-black text-amber-900 uppercase tracking-wider mb-1">
                  Panduan Konsolidasi Database (Sheet Working)
                </h3>
                <p className="text-xs text-amber-800 leading-relaxed font-semibold">
                  Gunakan tab ini untuk menggabungkan data ganda dengan
                  Kombinasi{" "}
                  <strong className="font-extrabold text-amber-950">
                    Checker + Channel + Hybrid + Lot No
                  </strong>{" "}
                  yang sama menjadi satu baris tunggal, di mana nilai stok
                  dipisahkan secara otomatis ke dalam kolom bulan yang tepat
                  (April - Maret).
                </p>
                <div className="mt-3 flex flex-col sm:flex-row gap-4 text-[11px] font-bold text-amber-900/90 border-t border-amber-200/40 pt-3">
                  <div className="flex items-center gap-1.5">
                    <span className="material-symbols-outlined text-[16px] text-amber-700">
                      visibility
                    </span>
                    <span>
                      <strong>Preview Konsolidasi:</strong> Simulasi visual di
                      layar (client-side) sebelum disimpan.
                    </span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <span className="material-symbols-outlined text-[16px] text-emerald-700">
                      auto_mode
                    </span>
                    <span>
                      <strong>Proses Konsolidasi:</strong>{" "}
                      Menyimpan/menggabungkan data secara permanen di Google
                      Sheet.
                    </span>
                  </div>
                </div>
                <div className="mt-3.5 bg-emerald-50/75 border border-emerald-200/50 rounded-xl p-3 text-[11px] font-semibold text-emerald-900 leading-relaxed shadow-inner flex items-center gap-2">
                  <span className="material-symbols-outlined text-[16px] text-emerald-600">
                    cloud_done
                  </span>
                  <span>
                    <strong>Koneksi Langsung Aktif:</strong> Sistem sekarang
                    terhubung langsung ke database Google Sheets melalui API
                    resmi. Anda tidak perlu lagi menyalin file Apps Script
                    manual atau melakukan langkah konfigurasi manual lainnya.
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* Stats Cards Row */}
          {(() => {
            const INDO_MONTHS = [
              "Jan",
              "Feb",
              "Mar",
              "Apr",
              "Mei",
              "Jun",
              "Jul",
              "Ags",
              "Sep",
              "Okt",
              "Nov",
              "Des",
            ];

            const getMonthIndexFromTimestamp = (timestamp: any): number => {
              if (!timestamp) return new Date().getMonth();
              if (timestamp instanceof Date) return timestamp.getMonth();

              const str = String(timestamp).trim();

              if (str.includes("/")) {
                const parts = str.split(/[\s/:]+/);
                if (parts.length >= 2) {
                  const mVal = parseInt(parts[1], 10);
                  if (!isNaN(mVal) && mVal >= 1 && mVal <= 12) {
                    return mVal - 1;
                  }
                  const months = [
                    "jan",
                    "feb",
                    "mar",
                    "apr",
                    "may",
                    "jun",
                    "jul",
                    "aug",
                    "sep",
                    "oct",
                    "nov",
                    "dec",
                  ];
                  const indMonths = [
                    "jan",
                    "feb",
                    "mar",
                    "apr",
                    "mei",
                    "jun",
                    "jul",
                    "agu",
                    "ags",
                    "sep",
                    "okt",
                    "nov",
                    "des",
                  ];
                  const lowerM = parts[1].toLowerCase();
                  let m = months.findIndex((name) => lowerM.startsWith(name));
                  if (m === -1)
                    m = indMonths.findIndex((name) => lowerM.startsWith(name));
                  if (m !== -1) return m;
                }
              }

              if (str.includes("-")) {
                const parts = str.split(/[\s\-:]+/);
                if (parts.length >= 2) {
                  const mVal = parseInt(parts[1], 10);
                  if (!isNaN(mVal) && mVal >= 1 && mVal <= 12) {
                    return mVal - 1;
                  }
                  const months = [
                    "jan",
                    "feb",
                    "mar",
                    "apr",
                    "may",
                    "jun",
                    "jul",
                    "aug",
                    "sep",
                    "oct",
                    "nov",
                    "dec",
                  ];
                  const indMonths = [
                    "jan",
                    "feb",
                    "mar",
                    "apr",
                    "mei",
                    "jun",
                    "jul",
                    "agu",
                    "ags",
                    "sep",
                    "okt",
                    "nov",
                    "des",
                  ];
                  const lowerM = parts[1].toLowerCase();
                  let m = months.findIndex((name) => lowerM.startsWith(name));
                  if (m === -1)
                    m = indMonths.findIndex((name) => lowerM.startsWith(name));
                  if (m !== -1) return m;
                }
              }

              const d = new Date(str);
              return !isNaN(d.getTime()) ? d.getMonth() : new Date().getMonth();
            };

            const groupedMap: Record<
              string,
              {
                id: string;
                checker: string;
                channel: string;
                hybrid: string;
                lot: string;
                shippingDate: string;
                expDate: string;
                inputs: Array<{
                  tanggalInput: string;
                  qty: number;
                  id: string;
                }>;
                monthlyQty: number[];
                totalQty: number;
              }
            > = {};

            const monthsKeys = [
              "jan",
              "feb",
              "mar",
              "apr",
              "mei",
              "jun",
              "jul",
              "ags",
              "sep",
              "okt",
              "nov",
              "des",
            ];

            (rawWorkingData && rawWorkingData.length > 0
              ? rawWorkingData
              : workingData
            ).forEach((item, index) => {
              const checker = item.user || item.pic || "Unknown";
              const channel =
                item.kiosk || item.channel || item.toko || "Unknown";
              const hybrid = item.hybrid || item.hybrids || "Unknown";
              const lot = item.lot || "Unknown";
              const key = `${checker}_${channel}_${hybrid}_${lot}`;

              let shippingDate =
                item.drDate ||
                item.shipping_date ||
                item.shippingDate ||
                item.dr_date ||
                "N/A";
              let expDate =
                item.expired ||
                item.exp_date ||
                item.expDate ||
                item.expired_date ||
                "N/A";

              const isValidVal = (v: any) =>
                v && v !== "N/A" && v !== "-" && String(v).trim() !== "";

              if (!isValidVal(shippingDate) || !isValidVal(expDate)) {
                const matchedDr = drSalesData.find(
                  (dr) =>
                    dr &&
                    dr.lot &&
                    cleanForMatch(dr.lot) === cleanForMatch(lot),
                );
                if (matchedDr) {
                  if (!isValidVal(shippingDate) && matchedDr.drDate) {
                    shippingDate = matchedDr.drDate;
                  }
                  if (!isValidVal(expDate) && matchedDr.expired) {
                    expDate = matchedDr.expired;
                  }
                }
              }

              // Determine monthly quantites for this item
              const itemMonthlyQty = Array(12).fill(0);
              let hasDbMonths = false;
              monthsKeys.forEach((mName, mIdx) => {
                if (
                  item[mName] !== undefined &&
                  item[mName] !== null &&
                  String(item[mName]).trim() !== ""
                ) {
                  itemMonthlyQty[mIdx] = Number(item[mName]) || 0;
                  hasDbMonths = true;
                }
              });

              const qtyVal = Number(item.stock) || Number(item.qty) || 0;
              if (!hasDbMonths) {
                const mIdx = getMonthIndexFromTimestamp(item.timestamp);
                itemMonthlyQty[mIdx] = qtyVal;
              }

              const itemTotal = itemMonthlyQty.reduce(
                (acc, val) => acc + val,
                0,
              );
              const tanggalInput = item.timestamp || "N/A";
              const inputId = item.id || `input_${index}`;

              if (!groupedMap[key]) {
                groupedMap[key] = {
                  id: `group_${index}`,
                  checker,
                  channel,
                  hybrid,
                  lot,
                  shippingDate,
                  expDate,
                  inputs: [],
                  monthlyQty: Array(12).fill(0),
                  totalQty: 0,
                };
              } else {
                const g = groupedMap[key];
                const isValidVal = (v: any) =>
                  v && v !== "N/A" && v !== "-" && String(v).trim() !== "";
                if (!isValidVal(g.shippingDate) && isValidVal(shippingDate)) {
                  g.shippingDate = shippingDate;
                }
                if (!isValidVal(g.expDate) && isValidVal(expDate)) {
                  g.expDate = expDate;
                }
              }

              // Save inputs details
              groupedMap[key].inputs.push({
                id: inputId,
                tanggalInput,
                qty: qtyVal,
              });

              // Sum monthly values
              for (let m = 0; m < 12; m++) {
                groupedMap[key].monthlyQty[m] += itemMonthlyQty[m];
              }
              groupedMap[key].totalQty += itemTotal;
            });

            const list = Object.values(groupedMap);

            const filteredList = list.filter((item) => {
              const q = tempSearchQuery.trim().toLowerCase();
              if (!q) return true;
              return (
                String(item.checker).toLowerCase().includes(q) ||
                String(item.channel).toLowerCase().includes(q) ||
                String(item.hybrid).toLowerCase().includes(q) ||
                String(item.lot).toLowerCase().includes(q)
              );
            });

            const uniqueCheckers = new Set(filteredList.map((i) => i.checker))
              .size;
            const uniqueChannels = new Set(filteredList.map((i) => i.channel))
              .size;
            const uniqueHybrids = new Set(filteredList.map((i) => i.hybrid))
              .size;
            const uniqueLots = new Set(filteredList.map((i) => i.lot)).size;

            // Handle Sorting
            const sortedList = [...filteredList].sort((a, b) => {
              const valA = String(a[tempSortBy] || "").toLowerCase();
              const valB = String(b[tempSortBy] || "").toLowerCase();
              if (valA < valB) return tempSortOrder === "asc" ? -1 : 1;
              if (valA > valB) return tempSortOrder === "asc" ? 1 : -1;
              return 0;
            });

            const toggleSort = (col: string) => {
              if (tempSortBy === col) {
                setTempSortOrder((prev) => (prev === "asc" ? "desc" : "asc"));
              } else {
                setTempSortBy(col);
                setTempSortOrder("asc");
              }
            };

            return (
              <>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
                  <div className="bg-white p-4 rounded-[20px] border border-[#edecff] shadow-sm flex items-center gap-3">
                    <div className="size-10 rounded-xl bg-indigo-50 flex items-center justify-center text-indigo-500 shrink-0">
                      <span className="material-symbols-outlined text-md">
                        assignment_ind
                      </span>
                    </div>
                    <div className="min-w-0">
                      <p className="text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider">
                        Unik Checker
                      </p>
                      <p className="text-sm font-bold text-[#181a2c] mt-0.5">
                        {uniqueCheckers}
                      </p>
                    </div>
                  </div>

                  <div className="bg-white p-4 rounded-[20px] border border-[#edecff] shadow-sm flex items-center gap-3">
                    <div className="size-10 rounded-xl bg-sky-50 flex items-center justify-center text-sky-500 shrink-0">
                      <span className="material-symbols-outlined text-md">
                        store
                      </span>
                    </div>
                    <div className="min-w-0">
                      <p className="text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider">
                        Unik Channel
                      </p>
                      <p className="text-sm font-bold text-[#181a2c] mt-0.5">
                        {uniqueChannels}
                      </p>
                    </div>
                  </div>

                  <div className="bg-white p-4 rounded-[20px] border border-[#edecff] shadow-sm flex items-center gap-3">
                    <div className="size-10 rounded-xl bg-emerald-50 flex items-center justify-center text-emerald-500 shrink-0">
                      <span className="material-symbols-outlined text-md">
                        category
                      </span>
                    </div>
                    <div className="min-w-0">
                      <p className="text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider">
                        Unik Hybrid
                      </p>
                      <p className="text-sm font-bold text-[#181a2c] mt-0.5">
                        {uniqueHybrids}
                      </p>
                    </div>
                  </div>

                  <div className="bg-white p-4 rounded-[20px] border border-[#edecff] shadow-sm flex items-center gap-3">
                    <div className="size-10 rounded-xl bg-purple-50 flex items-center justify-center text-purple-500 shrink-0">
                      <span className="material-symbols-outlined text-md">
                        layers
                      </span>
                    </div>
                    <div className="min-w-0">
                      <p className="text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider">
                        Unik Lot No
                      </p>
                      <p className="text-sm font-bold text-[#181a2c] mt-0.5">
                        {uniqueLots}
                      </p>
                    </div>
                  </div>
                </div>

                {/* Table Container */}
                <div className="bg-white rounded-[24px] border border-[#edecff] shadow-[0_12px_32px_rgba(21,75,226,0.12)] overflow-hidden">
                  {isTempProceeded && (
                    <div className="bg-amber-50 px-6 py-2.5 border-b border-amber-100 flex items-center gap-2 text-amber-800 text-[11px] font-bold">
                      <span className="material-symbols-outlined text-[16px] animate-pulse text-amber-600">
                        info
                      </span>
                      Geser tabel ke samping untuk melihat seluruh kolom bulan ↔
                    </div>
                  )}
                  <div className="overflow-x-auto min-w-full">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="border-b border-[#edecff] bg-[#fafbfe]/80">
                          {isTempProceeded ? (
                            <>
                              <th className="py-4 px-6 text-[10px] font-extrabold text-[#8E94B7] uppercase tracking-wider min-w-[130px]">
                                <button
                                  type="button"
                                  onClick={() => toggleSort("checker")}
                                  className="flex items-center gap-1.5 hover:text-primary transition-colors cursor-pointer text-left focus:outline-none"
                                >
                                  Checker
                                  <span className="material-symbols-outlined text-[14px]">
                                    {tempSortBy === "checker"
                                      ? tempSortOrder === "asc"
                                        ? "arrow_upward"
                                        : "arrow_downward"
                                      : "unfold_more"}
                                  </span>
                                </button>
                              </th>
                              <th className="py-4 px-6 text-[10px] font-extrabold text-[#8E94B7] uppercase tracking-wider min-w-[150px]">
                                <button
                                  type="button"
                                  onClick={() => toggleSort("channel")}
                                  className="flex items-center gap-1.5 hover:text-primary transition-colors cursor-pointer text-left focus:outline-none"
                                >
                                  Channel Partner
                                  <span className="material-symbols-outlined text-[14px]">
                                    {tempSortBy === "channel"
                                      ? tempSortOrder === "asc"
                                        ? "arrow_upward"
                                        : "arrow_downward"
                                      : "unfold_more"}
                                  </span>
                                </button>
                              </th>
                              <th className="py-4 px-6 text-[10px] font-extrabold text-[#8E94B7] uppercase tracking-wider min-w-[130px]">
                                <button
                                  type="button"
                                  onClick={() => toggleSort("hybrid")}
                                  className="flex items-center gap-1.5 hover:text-primary transition-colors cursor-pointer text-left focus:outline-none"
                                >
                                  Hybrid / Desc
                                  <span className="material-symbols-outlined text-[14px]">
                                    {tempSortBy === "hybrid"
                                      ? tempSortOrder === "asc"
                                        ? "arrow_upward"
                                        : "arrow_downward"
                                      : "unfold_more"}
                                  </span>
                                </button>
                              </th>
                              <th className="py-4 px-6 text-[10px] font-extrabold text-[#8E94B7] uppercase tracking-wider min-w-[110px]">
                                <button
                                  type="button"
                                  onClick={() => toggleSort("lot")}
                                  className="flex items-center gap-1.5 hover:text-primary transition-colors cursor-pointer text-left focus:outline-none"
                                >
                                  Lot No
                                  <span className="material-symbols-outlined text-[14px]">
                                    {tempSortBy === "lot"
                                      ? tempSortOrder === "asc"
                                        ? "arrow_upward"
                                        : "arrow_downward"
                                      : "unfold_more"}
                                  </span>
                                </button>
                              </th>
                              {INDO_MONTHS.map((m, mIdx) => (
                                <th
                                  key={mIdx}
                                  className="py-4 px-3 text-[10px] font-extrabold text-[#8E94B7] uppercase tracking-wider text-center min-w-[65px]"
                                >
                                  {m}
                                </th>
                              ))}
                              <th className="py-4 px-6 text-[10px] font-extrabold text-[#8E94B7] uppercase tracking-wider text-right min-w-[100px]">
                                Total Qty
                              </th>
                            </>
                          ) : (
                            <>
                              <th className="py-4 px-6 text-[10px] font-extrabold text-[#8E94B7] uppercase tracking-wider">
                                <button
                                  type="button"
                                  onClick={() => toggleSort("checker")}
                                  className="flex items-center gap-1.5 hover:text-primary transition-colors cursor-pointer text-left focus:outline-none"
                                >
                                  Checker
                                  <span className="material-symbols-outlined text-[14px]">
                                    {tempSortBy === "checker"
                                      ? tempSortOrder === "asc"
                                        ? "arrow_upward"
                                        : "arrow_downward"
                                      : "unfold_more"}
                                  </span>
                                </button>
                              </th>
                              <th className="py-4 px-6 text-[10px] font-extrabold text-[#8E94B7] uppercase tracking-wider">
                                <button
                                  type="button"
                                  onClick={() => toggleSort("channel")}
                                  className="flex items-center gap-1.5 hover:text-primary transition-colors cursor-pointer text-left focus:outline-none"
                                >
                                  Channel Partner
                                  <span className="material-symbols-outlined text-[14px]">
                                    {tempSortBy === "channel"
                                      ? tempSortOrder === "asc"
                                        ? "arrow_upward"
                                        : "arrow_downward"
                                      : "unfold_more"}
                                  </span>
                                </button>
                              </th>
                              <th className="py-4 px-6 text-[10px] font-extrabold text-[#8E94B7] uppercase tracking-wider">
                                <button
                                  type="button"
                                  onClick={() => toggleSort("hybrid")}
                                  className="flex items-center gap-1.5 hover:text-primary transition-colors cursor-pointer text-left focus:outline-none"
                                >
                                  Hybrid / Desc
                                  <span className="material-symbols-outlined text-[14px]">
                                    {tempSortBy === "hybrid"
                                      ? tempSortOrder === "asc"
                                        ? "arrow_upward"
                                        : "arrow_downward"
                                      : "unfold_more"}
                                  </span>
                                </button>
                              </th>
                              <th className="py-4 px-6 text-[10px] font-extrabold text-[#8E94B7] uppercase tracking-wider">
                                <button
                                  type="button"
                                  onClick={() => toggleSort("lot")}
                                  className="flex items-center gap-1.5 hover:text-primary transition-colors cursor-pointer text-left focus:outline-none"
                                >
                                  Lot No
                                  <span className="material-symbols-outlined text-[14px]">
                                    {tempSortBy === "lot"
                                      ? tempSortOrder === "asc"
                                        ? "arrow_upward"
                                        : "arrow_downward"
                                      : "unfold_more"}
                                  </span>
                                </button>
                              </th>
                            </>
                          )}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#f2f1ff] text-xs font-semibold text-[#181a2c]">
                        {sortedList.length > 0 ? (
                          sortedList.map((item, idx) => {
                            const isExpanded = expandedTempRowId === item.id;
                            if (isTempProceeded) {
                              return (
                                <tr
                                  key={item.id}
                                  className="hover:bg-[#fbfbfb] transition-colors"
                                >
                                  <td className="py-4 px-6 select-text">
                                    <div className="flex items-center gap-2.5">
                                      <div className="size-7 rounded-lg bg-indigo-50 text-indigo-600 font-bold text-[10px] flex items-center justify-center uppercase shrink-0 shadow-sm animate-in fade-in zoom-in-95 duration-150">
                                        {String(item.checker).substring(0, 2)}
                                      </div>
                                      <span className="truncate max-w-[150px] font-bold text-[#181a2c]">
                                        {item.checker}
                                      </span>
                                    </div>
                                  </td>
                                  <td className="py-4 px-6 text-[#5e617d] truncate max-w-[180px] select-text">
                                    {item.channel}
                                  </td>
                                  <td className="py-4 px-6 select-text">
                                    <span className="bg-primary/5 text-primary text-[10.5px] px-3 py-1 rounded-lg font-bold inline-block leading-none border border-primary/10">
                                      {item.hybrid}
                                    </span>
                                  </td>
                                  <td className="py-4 px-6 select-text">
                                    <span className="font-mono text-[11px] font-extrabold px-2.5 py-1 bg-slate-50 border border-slate-100 rounded text-slate-700 tracking-wide font-bold">
                                      {item.lot}
                                    </span>
                                  </td>
                                  {item.monthlyQty.map((val, mIdx) => (
                                    <td
                                      key={mIdx}
                                      className="py-4 px-3 text-center"
                                    >
                                      {val > 0 ? (
                                        <span className="bg-primary/5 text-primary text-[10px] px-2 py-0.5 rounded font-extrabold inline-block scale-100 hover:scale-105 duration-100 border border-primary/5">
                                          {val.toLocaleString()}
                                        </span>
                                      ) : (
                                        <span className="text-[#c1c4db] font-normal">
                                          -
                                        </span>
                                      )}
                                    </td>
                                  ))}
                                  <td className="py-4 px-6 text-right font-extrabold text-primary select-all">
                                    {item.totalQty.toLocaleString()} Kg
                                  </td>
                                </tr>
                              );
                            }

                            // Otherwise, normal row with expandability
                            return (
                              <React.Fragment key={item.id}>
                                <tr
                                  onClick={() =>
                                    setExpandedTempRowId(
                                      isExpanded ? null : item.id,
                                    )
                                  }
                                  className={`hover:bg-[#fbfbfb] border-l-4 transition-all cursor-pointer select-none ${
                                    isExpanded
                                      ? "bg-primary/[0.02] border-primary"
                                      : "border-transparent"
                                  }`}
                                >
                                  <td className="py-4 px-6 select-text">
                                    <div className="flex items-center gap-2.5">
                                      <span
                                        className="material-symbols-outlined text-[20px] text-[#8E94B7] transition-transform duration-300 shrink-0 select-none"
                                        style={{
                                          transform: isExpanded
                                            ? "rotate(90deg)"
                                            : "none",
                                        }}
                                      >
                                        chevron_right
                                      </span>
                                      <div className="size-7 rounded-lg bg-indigo-50 text-indigo-600 font-bold text-[10px] flex items-center justify-center uppercase shrink-0 shadow-sm relative">
                                        {String(item.checker).substring(0, 2)}
                                        {item.inputs.length > 1 && (
                                          <span className="absolute -top-1.5 -right-1.5 bg-primary text-white text-[8px] h-3.5 min-w-3.5 px-0.5 rounded-full flex items-center justify-center font-bold font-sans">
                                            {item.inputs.length}
                                          </span>
                                        )}
                                      </div>
                                      <span className="truncate max-w-[150px] font-bold text-[#181a2c]">
                                        {item.checker}
                                      </span>
                                    </div>
                                  </td>
                                  <td className="py-4 px-6 text-[#5e617d] truncate max-w-[180px] select-text">
                                    {item.channel}
                                  </td>
                                  <td className="py-4 px-6 select-text">
                                    <span className="bg-primary/5 text-primary text-[10.5px] px-3 py-1 rounded-lg font-bold inline-block leading-none border border-primary/10">
                                      {item.hybrid}
                                    </span>
                                  </td>
                                  <td className="py-4 px-6 select-text">
                                    <span className="font-mono text-[11px] font-extrabold px-2.5 py-1 bg-slate-50 border border-slate-100 rounded text-slate-700 tracking-wide font-bold">
                                      {item.lot}
                                    </span>
                                  </td>
                                </tr>
                                {isExpanded && (
                                  <tr className="bg-[#fafbfe]/40">
                                    <td colSpan={4} className="py-3 px-6">
                                      <div className="p-5 bg-white rounded-2xl border border-[#edecff] shadow-sm animate-in slide-in-from-top-3 duration-200">
                                        <h4 className="text-[10.5px] font-extrabold text-primary uppercase tracking-widest mb-3.5 flex items-center gap-1.5">
                                          <span className="material-symbols-outlined text-[16px]">
                                            calendar_today
                                          </span>{" "}
                                          Detail Tanggal Input & Qty
                                        </h4>

                                        {/* Shipping Date & Exp Date Detail Cards */}
                                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
                                          <div className="flex items-center gap-3 bg-slate-50 border border-[#edecff] rounded-xl p-3">
                                            <div className="size-8 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center shrink-0">
                                              <span className="material-symbols-outlined text-[18px]">
                                                local_shipping
                                              </span>
                                            </div>
                                            <div className="min-w-0">
                                              <p className="text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider">
                                                Shipping Date
                                              </p>
                                              <p className="text-xs font-black text-[#181a2c] mt-0.5">
                                                {item.shippingDate || "N/A"}
                                              </p>
                                            </div>
                                          </div>
                                          <div className="flex items-center gap-3 bg-slate-50 border border-[#edecff] rounded-xl p-3">
                                            <div className="size-8 rounded-lg bg-rose-50 text-rose-600 flex items-center justify-center shrink-0">
                                              <span className="material-symbols-outlined text-[18px]">
                                                event_busy
                                              </span>
                                            </div>
                                            <div className="min-w-0">
                                              <p className="text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider">
                                                Exp Date
                                              </p>
                                              <p className="text-xs font-black text-rose-700 mt-0.5">
                                                {item.expDate || "N/A"}
                                              </p>
                                            </div>
                                          </div>
                                        </div>

                                        <h4 className="text-[10.5px] font-extrabold text-primary uppercase tracking-widest mb-3.5 flex items-center gap-1.5">
                                          <span className="material-symbols-outlined text-[16px]">
                                            calendar_today
                                          </span>{" "}
                                          Detail Tanggal Input & Qty
                                        </h4>
                                        <div className="overflow-hidden rounded-xl border border-[#edecff] bg-white">
                                          <table className="w-full text-left border-collapse">
                                            <thead>
                                              <tr className="bg-slate-50 border-b border-[#edecff] text-[9.5px] font-extrabold text-[#8E94B7] uppercase tracking-wider">
                                                <th className="py-2.5 px-4 w-12 text-center">
                                                  No
                                                </th>
                                                <th className="py-2.5 px-4">
                                                  Tanggal Input
                                                </th>
                                                <th className="py-2.5 px-4 text-right">
                                                  Stok / Qty (Kg)
                                                </th>
                                              </tr>
                                            </thead>
                                            <tbody className="divide-y divide-[#f2f1ff] text-xs font-semibold text-[#181a2c]">
                                              {item.inputs.map(
                                                (input, inputIdx) => (
                                                  <tr
                                                    key={input.id || inputIdx}
                                                    className="hover:bg-slate-50/50 transition-colors"
                                                  >
                                                    <td className="py-3 px-4 text-[#8E94B7] text-center font-mono text-[11px]">
                                                      {inputIdx + 1}
                                                    </td>
                                                    <td className="py-3 px-4 font-mono text-[#5e617d] select-all">
                                                      {input.tanggalInput}
                                                    </td>
                                                    <td className="py-3 px-4 font-extrabold text-right text-primary select-all">
                                                      {input.qty.toLocaleString()}{" "}
                                                      Kg
                                                    </td>
                                                  </tr>
                                                ),
                                              )}
                                            </tbody>
                                            <tfoot>
                                              <tr className="bg-slate-50/30 border-t border-[#edecff] font-bold text-[#181a2c] text-xs">
                                                <td
                                                  colSpan={2}
                                                  className="py-2.5 px-4 text-[10px] uppercase font-extrabold text-[#8E94B7] text-left"
                                                >
                                                  Total Akumulasi
                                                </td>
                                                <td className="py-2.5 px-4 text-right font-extrabold text-primary select-all text-sm">
                                                  {item.totalQty.toLocaleString()}{" "}
                                                  Kg
                                                </td>
                                              </tr>
                                            </tfoot>
                                          </table>
                                        </div>
                                      </div>
                                    </td>
                                  </tr>
                                )}
                              </React.Fragment>
                            );
                          })
                        ) : (
                          <tr>
                            <td
                              colSpan={isTempProceeded ? 17 : 4}
                              className="py-16 text-center text-[#8E94B7] font-medium"
                            >
                              Tidak ada data yang cocok dengan kriteria
                              pencarian.
                            </td>
                          </tr>
                        )}
                      </tbody>
                      {isTempProceeded &&
                        sortedList.length > 0 &&
                        (() => {
                          const monthColumnTotals = Array(12).fill(0);
                          sortedList.forEach((item) => {
                            item.monthlyQty.forEach((val, idx) => {
                              monthColumnTotals[idx] += val;
                            });
                          });
                          const granTotal = sortedList.reduce(
                            (sum, item) => sum + item.totalQty,
                            0,
                          );
                          return (
                            <tfoot>
                              <tr className="bg-[#fafbfe]/80 border-t border-[#edecff] font-bold text-[#181a2c] text-xs">
                                <td
                                  colSpan={4}
                                  className="py-3 px-6 text-[10px] uppercase font-extrabold text-[#8E94B7] text-left"
                                >
                                  Total Kolom
                                </td>
                                {monthColumnTotals.map((tot, idx) => (
                                  <td
                                    key={idx}
                                    className="py-3 px-3 text-center font-extrabold text-[#154be2] text-[11px]"
                                  >
                                    {tot > 0 ? tot.toLocaleString() : "-"}
                                  </td>
                                ))}
                                <td className="py-3 px-6 text-right font-extrabold text-primary select-all text-xs">
                                  {granTotal.toLocaleString()} Kg
                                </td>
                              </tr>
                            </tfoot>
                          );
                        })()}
                    </table>
                  </div>
                  {sortedList.length > 0 && (
                    <div className="py-3.5 px-6 border-t border-[#edecff] bg-slate-50/30 text-[10px] text-[#8E94B7] font-bold text-right uppercase tracking-wider">
                      Total Data: {sortedList.length} Baris
                    </div>
                  )}
                </div>
              </>
            );
          })()}
        </div>
      )}



      {activeTab === "propose" && (
        <div className="animate-in fade-in slide-in-from-right-4 duration-300">
          {/* Header */}
          <div className="mb-6 ml-1 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="size-10 rounded-2xl bg-gradient-to-br from-[#154be2]/20 to-cyan-400/20 flex items-center justify-center shrink-0 border border-[#154be2]/20 shadow-inner">
                <span className="material-symbols-outlined text-[#154be2] text-[20px]">
                  rate_review
                </span>
              </div>
              <div>
                <h1 className="text-lg font-bold text-[#181a2c] tracking-tight">
                  Propose Activity
                </h1>
                <p className="text-[#8E94B7] text-[11px] font-semibold tracking-wide mt-0.5">
                  Formulir pengajuan kegiatan promosi dan demo produk RADAR DG
                </p>
              </div>
            </div>
          </div>

          {/* Success Notification */}
          {proposalSuccessMsg && (
            <div className="mb-5 bg-emerald-50 border border-emerald-200 text-emerald-800 rounded-2xl px-5 py-3 text-xs font-bold flex items-center gap-2.5 shadow-sm animate-in fade-in slide-in-from-top-2 duration-300">
              <span className="material-symbols-outlined text-emerald-600 text-[20px]">
                check_circle
              </span>
              <div className="flex-1">{proposalSuccessMsg}</div>
              <button 
                onClick={() => setProposalSuccessMsg("")} 
                className="text-emerald-500 hover:text-emerald-700 font-bold text-sm cursor-pointer"
              >
                ×
              </button>
            </div>
          )}

          {/* Main Content Grid */}
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
            {/* Left Panel: Form Input */}
            <div className="lg:col-span-7 bg-white rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.12)] border border-[#154be2]/5 p-6 space-y-6">
              
              {/* BS Name Selector - CHANGED TO PICKLIST */}
              <div>
                <label className="block text-xs font-bold text-[#181a2c] mb-2.5 uppercase tracking-wider">
                  Nama BS (Business Solution)
                </label>
                <div className="relative">
                  <span className="absolute inset-y-0 left-0 flex items-center pl-3.5 pointer-events-none text-slate-400">
                    <span className="material-symbols-outlined text-[18px]">person</span>
                  </span>
                  <select
                    value={proposeBs}
                    onChange={(e) => {
                      setProposeBs(e.target.value);
                    }}
                    className="w-full bg-slate-50/50 hover:bg-slate-50 text-slate-800 text-xs font-semibold pl-10 pr-10 py-2.5 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-[#154be2]/20 focus:border-[#154be2] transition-colors cursor-pointer appearance-none"
                  >
                    {bsEmployeesList.map((name) => (
                      <option key={name} value={name}>{name}</option>
                    ))}
                  </select>
                  <span className="absolute inset-y-0 right-0 flex items-center pr-3.5 pointer-events-none text-slate-400">
                    <span className="material-symbols-outlined text-[16px]">expand_more</span>
                  </span>
                </div>
              </div>

              {/* Activity Category */}
              <div>
                <label className="block text-xs font-bold text-[#181a2c] mb-2.5 uppercase tracking-wider">
                  Activity Category
                </label>
                <div className="flex bg-slate-100 p-1 rounded-xl">
                  {["Regular", "AdHoc"].map((cat) => (
                    <button
                      key={cat}
                      type="button"
                      onClick={() => {
                        setProposeCategory(cat);
                        setProposeActivity(cat === "Regular" ? "Farmer meeting" : "AIC");
                      }}
                      className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${proposeCategory === cat ? "bg-white text-[#154be2] shadow-sm" : "text-slate-500 hover:text-slate-700"}`}
                    >
                      {cat}
                    </button>
                  ))}
                </div>
              </div>

              {/* Activity Type */}
              <div>
                <label className="block text-xs font-bold text-[#181a2c] mb-2.5 uppercase tracking-wider">
                  Activity Type
                </label>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                  {(proposeCategory === "Regular" ? ["Farmer meeting", "Farmer field day", "One day promo", "Special field trip"] : ["AIC", "Expo", "Caravan", "Retailer Meeting"]).map((act) => {
                    const isSelected = proposeActivity === act;
                    const budgets: Record<string, {actual: number, remaining: number}> = {
                      "Farmer meeting": { actual: 5000000, remaining: 12000000 },
                      "Farmer field day": { actual: 8000000, remaining: 20000000 },
                      "One day promo": { actual: 2000000, remaining: 5000000 },
                      "Special field trip": { actual: 15000000, remaining: 35000000 },
                      "AIC": { actual: 3000000, remaining: 10000000 },
                      "Expo": { actual: 25000000, remaining: 50000000 },
                      "Caravan": { actual: 10000000, remaining: 25000000 },
                      "Retailer Meeting": { actual: 4000000, remaining: 8000000 },
                    };
                    const b = budgets[act] || { actual: 0, remaining: 0 };
                    
                    return (
                      <button
                        key={act}
                        type="button"
                        onClick={() => setProposeActivity(act)}
                        className={`flex flex-col items-start p-3 rounded-xl border text-left transition-all duration-200 ${
                          isSelected
                            ? "bg-[#154be2]/5 border-[#154be2] shadow-sm ring-1 ring-[#154be2]"
                            : "border-slate-200 bg-white hover:bg-slate-50"
                        }`}
                      >
                        <span className={`text-[10px] sm:text-[11px] font-bold mb-2 ${isSelected ? 'text-[#154be2]' : 'text-slate-700'}`}>{act}</span>
                        <div className="mt-auto w-full">
                          <div className="flex justify-between items-center text-[9px] mb-0.5">
                            <span className="text-slate-400 font-medium">Actual:</span>
                            <span className="font-semibold text-slate-700">Rp{(b.actual/1000000).toFixed(1)}M</span>
                          </div>
                          <div className="flex justify-between items-center text-[9px]">
                            <span className="text-slate-400 font-medium">Remaining:</span>
                            <span className="font-bold text-emerald-600">Rp{(b.remaining/1000000).toFixed(1)}M</span>
                          </div>
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Planning Data List */}
              <div className="mt-6">
                <div className="flex justify-between items-center mb-2.5">
                  <label className="text-xs font-bold text-[#181a2c] uppercase tracking-wider">
                    Planning Data
                  </label>
                  <div className="flex items-center space-x-2">
                    <span className="text-[10px] text-slate-500 font-semibold">Status:</span>
                    <select
                      value={planningStatusFilter}
                      onChange={(e) => setPlanningStatusFilter(e.target.value)}
                      className="text-[10px] font-bold border border-slate-200 rounded-lg text-slate-700 bg-white focus:outline-none focus:border-[#154be2] px-2 py-1 shadow-sm"
                    >
                      <option value="All">All</option>
                      <option value="Remaining">Remaining</option>
                      <option value="Complete">Complete</option>
                    </select>
                  </div>
                </div>
                <div className="border border-slate-200 rounded-xl overflow-hidden bg-white">
                  <div className="overflow-x-auto">
                    {(() => {
                      const rawData = [
                        { id: "L1", district: "Malang", subDistrict: "Waru", budget: 1500000, month: "Agustus", status: "Remaining" },
                        { id: "L2", district: "Malang", subDistrict: "Singosari", budget: 2000000, month: "September", status: "Complete" },
                        { id: "L3", district: "Pasuruan", subDistrict: "Bangil", budget: 1200000, month: "Oktober", status: "Remaining" },
                        { id: "L4", district: "Pasuruan", subDistrict: "Pandaan", budget: 1800000, month: "Agustus", status: "Complete" },
                        { id: "L5", district: "Batu", subDistrict: "Bumiaji", budget: 1500000, month: "November", status: "Remaining" },
                      ];
                      
                      const uniqueDistricts = Array.from(new Set(rawData.map(r => r.district))).sort();
                      const uniqueSubDistricts = Array.from(new Set(rawData.map(r => r.subDistrict))).sort();
                      const uniqueMonths = Array.from(new Set(rawData.map(r => r.month))).sort();
                      
                      const filteredData = rawData.filter(row => 
                        (planningFilterDistrict === "" || row.district === planningFilterDistrict) &&
                        (planningFilterSubDistrict === "" || row.subDistrict === planningFilterSubDistrict) &&
                        (planningFilterMonth === "" || row.month === planningFilterMonth) &&
                        (planningStatusFilter === "All" || row.status === planningStatusFilter)
                      );

                      return (
                        <table className="w-full text-left border-collapse min-w-[500px]">
                          <thead>
                            <tr className="bg-slate-50">
                              <th className="px-3 py-2 text-[9px] font-extrabold text-slate-500 uppercase tracking-wider border-b border-slate-200">
                                District
                                <select value={planningFilterDistrict} onChange={e => setPlanningFilterDistrict(e.target.value)} className="mt-1 block w-full px-1.5 py-1 text-[9px] font-normal border border-slate-200 rounded text-slate-800 bg-white focus:outline-none focus:border-[#154be2]">
                                  <option value="">All</option>
                                  {uniqueDistricts.map(d => <option key={d} value={d}>{d}</option>)}
                                </select>
                              </th>
                              <th className="px-3 py-2 text-[9px] font-extrabold text-slate-500 uppercase tracking-wider border-b border-slate-200">
                                Sub District
                                <select value={planningFilterSubDistrict} onChange={e => setPlanningFilterSubDistrict(e.target.value)} className="mt-1 block w-full px-1.5 py-1 text-[9px] font-normal border border-slate-200 rounded text-slate-800 bg-white focus:outline-none focus:border-[#154be2]">
                                  <option value="">All</option>
                                  {uniqueSubDistricts.map(d => <option key={d} value={d}>{d}</option>)}
                                </select>
                              </th>
                              <th className="px-3 py-2 text-[9px] font-extrabold text-slate-500 uppercase tracking-wider border-b border-slate-200 align-top">Budget</th>
                              <th className="px-3 py-2 text-[9px] font-extrabold text-slate-500 uppercase tracking-wider border-b border-slate-200">
                                Month
                                <select value={planningFilterMonth} onChange={e => setPlanningFilterMonth(e.target.value)} className="mt-1 block w-full px-1.5 py-1 text-[9px] font-normal border border-slate-200 rounded text-slate-800 bg-white focus:outline-none focus:border-[#154be2]">
                                  <option value="">All</option>
                                  {uniqueMonths.map(m => <option key={m} value={m}>{m}</option>)}
                                </select>
                              </th>
                              <th className="px-3 py-2 text-[9px] font-extrabold text-slate-500 uppercase tracking-wider border-b border-slate-200 text-right align-top">Aksi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100">
                            {filteredData.map((row) => {
                              const isComplete = row.status === "Complete";
                              const isProcessing = processingRows[row.id];
                              
                              let btnClass = "bg-[#154be2] hover:bg-[#154be2]/90 text-white";
                              let btnText = "Propose";
                              let disabled = false;
                              
                              if (isComplete) {
                                btnClass = "bg-slate-300 text-slate-500 cursor-not-allowed";
                                disabled = true;
                              } else if (isProcessing) {
                                btnClass = "bg-yellow-400 text-yellow-900 cursor-not-allowed";
                                btnText = "Process";
                                disabled = true;
                              }

                              return (
                                <tr key={row.id} className="hover:bg-slate-50/50 transition-colors">
                                  <td className="px-3 py-2 text-[10px] font-bold text-slate-700">{row.district}</td>
                                  <td className="px-3 py-2 text-[10px] font-semibold text-slate-600">{row.subDistrict}</td>
                                  <td className="px-3 py-2 text-[10px] font-bold text-slate-800">Rp {row.budget.toLocaleString()}</td>
                                  <td className="px-3 py-2 text-[10px] font-medium text-slate-500">{row.month}</td>
                                  <td className="px-3 py-2 text-right">
                                    <div className="flex flex-col sm:flex-row items-end sm:items-center justify-end gap-2">
                                      <span className={`text-[9px] font-bold px-2 py-0.5 rounded-full ${isComplete ? 'bg-green-100 text-green-700' : 'bg-blue-100 text-blue-700'}`}>
                                        {row.status}
                                      </span>
                                      <button
                                        type="button"
                                        disabled={disabled}
                                        onClick={() => {
                                          setProcessingRows(prev => ({ ...prev, [row.id]: true }));
                                          
                                          const bsMap = { "Lionel Messi": "lm", "Ronaldo": "ro", "Mbappe": "mb", "Yamal": "ya" };
                                          const actMap = { "Farmer meeting": "fm", "Farmer field day": "ffd", "One day promo": "odp", "Special field trip": "sft", "AIC": "aic", "Expo": "exp", "Caravan": "crv", "Retailer Meeting": "rm" };
                                          const bsCode = bsMap[proposeBs as keyof typeof bsMap] || proposeBs.substring(0, 2).toLowerCase();
                                          const actCode = actMap[proposeActivity as keyof typeof actMap] || "act";
                                          const monthMap = { "Agustus": "08", "September": "09", "Oktober": "10", "November": "11" };
                                          const monthCode = monthMap[row.month as keyof typeof monthMap] || "08";
                                          const yearCode = "26";
                                          const uniqueSuffix = String(Math.floor(Math.random() * 900) + 100);
                                          const projectNo = `${bsCode}/${actCode}/${monthCode}/${yearCode}-${uniqueSuffix}`;
                                          
                                          const newProj = {
                                            id: (Date.now() + Math.random()).toString(),
                                            projectNo,
                                            bs: proposeBs,
                                            category: proposeCategory,
                                            activity: proposeActivity,
                                            district: row.district,
                                            subDistrict: row.subDistrict,
                                            budget: row.budget,
                                            month: row.month,
                                            farmerReach: "",
                                            hybrids: "",
                                          };
                                          setGeneratedProjects((prev) => [...prev, newProj]);
                                        }}
                                        className={`px-3 py-1.5 rounded-lg text-[10px] font-bold transition-all shadow-sm ${!disabled ? 'active:scale-95' : ''} ${btnClass}`}
                                      >
                                        {btnText}
                                      </button>
                                    </div>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      );
                    })()}
                  </div>
                </div>
              </div>
            </div>

            {/* Right Panel: Output & Submission Card */}
            <div className="lg:col-span-5 h-full space-y-4">
              {generatedProjects.length > 0 ? (
                <div className="bg-white rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.12)] border border-[#154be2]/5 p-6 space-y-6 animate-in fade-in zoom-in-95 duration-300">
                  <div className="flex flex-col items-center text-center">
                    <div className="size-12 rounded-full bg-[#154be2]/10 flex items-center justify-center text-[#154be2] mb-3">
                      <span className="material-symbols-outlined text-[24px]">qr_code_2</span>
                    </div>
                    <h3 className="text-sm font-bold text-[#181a2c]">Generated Project Numbers</h3>
                    <p className="text-[10px] font-semibold text-[#8E94B7] mt-0.5">
                      Berhasil mengenerate {generatedProjects.length} nomor proyek. Tekan Submit untuk menyimpan.
                    </p>
                  </div>

                  {/* List of Generated Projects */}
                  <div className="overflow-x-auto max-h-[300px] overflow-y-auto">
                    <table className="w-full text-left border-collapse min-w-[500px]">
                      <thead>
                        <tr className="bg-[#fbfaff]">
                          <th className="px-4 py-2 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9] sticky top-0 bg-[#fbfaff] z-10">Project No</th>
                          <th className="px-4 py-2 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9] sticky top-0 bg-[#fbfaff] z-10">BS Name</th>
                          <th className="px-4 py-2 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9] sticky top-0 bg-[#fbfaff] z-10">Activity</th>
                          <th className="px-4 py-2 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9] sticky top-0 bg-[#fbfaff] z-10">Location</th>
                          <th className="px-4 py-2 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9] sticky top-0 bg-[#fbfaff] z-10">Budget</th>
                          <th className="px-4 py-2 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9] sticky top-0 bg-[#fbfaff] z-10">Month</th>
                          <th className="px-4 py-2 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9] text-right sticky top-0 bg-[#fbfaff] z-10">Aksi</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#f1f5f9]">
                        {generatedProjects.map((proj, idx) => (
                          <tr key={proj.id} className="hover:bg-slate-50/50 transition-colors">
                            <td className="px-4 py-2.5">
                              <span className="font-mono text-[10px] font-bold text-[#154be2] tracking-wider select-all">
                                {proj.projectNo}
                              </span>
                            </td>
                            <td className="px-4 py-2.5 text-[10px] font-bold text-[#181a2c]">
                              {proj.bs}
                            </td>
                            <td className="px-4 py-2.5 text-[10px] font-semibold text-[#8E94B7]">
                              {proj.activity}
                            </td>
                            <td className="px-4 py-2.5 text-[10px] font-semibold text-[#8E94B7]">
                              {proj.district}, {proj.subDistrict}
                            </td>
                            <td className="px-4 py-2.5 text-[10px] font-semibold text-[#8E94B7]">
                              Rp {proj.budget?.toLocaleString()}
                            </td>
                            <td className="px-4 py-2.5 text-[10px] font-semibold text-[#8E94B7]">
                              {proj.month}
                            </td>
                            <td className="px-4 py-2.5 text-right">
                              <button
                                type="button"
                                onClick={() => setGeneratedProjects(generatedProjects.filter((p) => p.id !== proj.id))}
                                className="text-slate-400 hover:text-red-500 cursor-pointer transition-colors p-1 bg-white rounded shadow-sm border border-slate-200 hover:border-red-200 flex items-center justify-center ml-auto"
                                title="Hapus dari antrean"
                              >
                                <span className="material-symbols-outlined text-[14px]">delete</span>
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>


                  {/* Submit CTA */}
                  <button
                    type="button"
                    disabled={isSubmittingProposal}
                    onClick={() => {
                      if (generatedProjects.length === 0) return;
                      setIsSubmittingProposal(true);
                      
                      setTimeout(() => {
                        const withTimestamps = generatedProjects.map((p) => ({
                          ...p,
                          createdAt: new Date().toLocaleString("id-ID"),
                        }));

                        const updatedList = [...withTimestamps, ...proposalsList];
                        setProposalsList(updatedList);
                        try {
                          localStorage.setItem("radar_dg_proposals", JSON.stringify(updatedList));
                        } catch (e) {
                          console.error("Failed to save proposal list", e);
                        }

                        setIsSubmittingProposal(false);
                        setProposalSuccessMsg(`Sukses! Berhasil mengirim ${generatedProjects.length} proyek ke database (dummy).`);
                        setGeneratedProjects([]); // Clear staging area
                      }, 1000);
                    }}
                    className="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-emerald-400 text-white py-3 rounded-xl font-bold text-xs transition-all shadow-sm active:scale-98 cursor-pointer"
                  >
                    {isSubmittingProposal ? (
                      <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin mr-1" />
                    ) : (
                      <span className="material-symbols-outlined text-[18px]">cloud_upload</span>
                    )}
                    {isSubmittingProposal ? "Submitting..." : "Submit to Database"}
                  </button>
                </div>
              ) : (
                <div className="bg-white rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.12)] border border-[#154be2]/5 p-8 text-center flex flex-col items-center justify-center min-h-[360px]">
                  <div className="size-16 rounded-full bg-slate-50 flex items-center justify-center text-slate-300 mb-4 border border-dashed border-slate-200">
                    <span className="material-symbols-outlined text-[32px]">post_add</span>
                  </div>
                  <h4 className="text-xs font-bold text-[#181a2c] uppercase tracking-wider mb-1.5">Proposal Generation</h4>
                  <p className="text-[10px] text-[#8E94B7] font-semibold leading-relaxed max-w-xs">
                    Isi seluruh formulir pengajuan di sebelah kiri secara lengkap, lalu tekan tombol <strong className="text-primary">Propose Activity</strong> untuk mengenerate nomor nomor proyek otomatis. Anda bisa mengenerate beberapa nomor proyek sekaligus!
                  </p>
                </div>
              )}
            </div>
          </div>

          {/* History Section */}
          <div className="bg-white rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.12)] border border-[#154be2]/5 overflow-hidden mt-6">
            <div className="p-5 border-b border-[#f1f5f9] flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
              <div className="flex items-center gap-2">
                <div className="size-8 rounded-lg bg-indigo-50 flex items-center justify-center">
                  <span className="material-symbols-outlined text-[#154be2] text-[18px]">
                    history
                  </span>
                </div>
                <div>
                  <h3 className="text-xs font-bold text-[#181a2c] uppercase tracking-wider">Proposal History</h3>
                  <p className="text-[9.5px] font-semibold text-[#8E94B7] mt-0.5">Daftar project pengajuan yang berhasil dikirim</p>
                </div>
              </div>

              {proposalsList.length > 0 && (
                <button
                  onClick={() => {
                    if (window.confirm("Apakah Anda yakin ingin menghapus semua history dummy?")) {
                      setProposalsList([]);
                      try {
                        localStorage.removeItem("radar_dg_proposals");
                      } catch (e) {}
                    }
                  }}
                  className="text-red-500 hover:text-red-600 hover:bg-red-50 px-3 py-1.5 rounded-lg text-[10px] font-bold transition-all flex items-center gap-1 border border-red-200/50 cursor-pointer"
                >
                  <span className="material-symbols-outlined text-[14px]">delete_sweep</span>
                  Hapus Semua Dummy
                </button>
              )}
            </div>

            <div className="px-5 py-2.5 bg-slate-50 border-b border-[#f1f5f9] flex items-center gap-1.5 text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider">
              <span className="material-symbols-outlined text-[14px] text-[#154be2]">info</span>
              <span>Klik pada nomor proyek untuk melakukan drill down (melihat rincian detail kegiatan)</span>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse min-w-[750px]">
                <thead>
                  <tr className="bg-[#fbfaff]">
                    <th className="px-5 py-3 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Project No</th>
                    <th className="px-5 py-3 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">BS Name</th>
                    <th className="px-5 py-3 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Activity</th>
                    <th className="px-5 py-3 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Location</th>
                    <th className="px-5 py-3 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Budget</th>
                    <th className="px-5 py-3 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Month</th>
                    <th className="px-5 py-3 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Submitted At</th>
                    <th className="px-5 py-3 text-[9px] font-extrabold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9] text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#f1f5f9]">
                  {proposalsList.length > 0 ? (
                    proposalsList.map((proposal) => (
                      <React.Fragment key={proposal.id}>
                        <tr className="hover:bg-slate-50/50 transition-colors">
                          <td className="px-5 py-3.5">
                            <button
                              type="button"
                              onClick={() => {
                                setExpandedProposalId(expandedProposalId === proposal.id ? null : proposal.id);
                              }}
                              className="font-mono text-[11px] font-extrabold text-[#154be2] bg-[#154be2]/5 hover:bg-[#154be2]/10 px-2.5 py-1 rounded border border-[#154be2]/15 flex items-center gap-1.5 cursor-pointer transition-all"
                              title="Klik untuk melihat detail proyek"
                            >
                              <span>{proposal.projectNo}</span>
                              <span className="material-symbols-outlined text-[14px] leading-none text-[#154be2]/70">
                                {expandedProposalId === proposal.id ? "keyboard_arrow_up" : "keyboard_arrow_down"}
                              </span>
                            </button>
                          </td>
                          <td className="px-5 py-3.5 text-xs font-semibold text-[#181a2c]">
                            {proposal.bs}
                          </td>
                          <td className="px-5 py-3.5 text-xs font-semibold text-[#181a2c]">
                            {proposal.activity}
                          </td>
                          <td className="px-5 py-3.5 text-xs font-bold text-slate-700">
                            {proposal.district}, {proposal.subDistrict}
                          </td>
                          <td className="px-5 py-3.5 text-xs font-medium text-slate-700">
                            Rp {proposal.budget?.toLocaleString()}
                          </td>
                          <td className="px-5 py-3.5 text-xs font-medium text-slate-500">
                            {proposal.month}
                          </td>
                          <td className="px-5 py-3.5 text-[10px] font-semibold text-slate-400">
                            {proposal.createdAt}
                          </td>
                          <td className="px-5 py-3.5 text-right">
                            <button
                              onClick={() => {
                                const newList = proposalsList.filter((item) => item.id !== proposal.id);
                                setProposalsList(newList);
                                try {
                                  localStorage.setItem("radar_dg_proposals", JSON.stringify(newList));
                                } catch (e) {}
                              }}
                              className="text-red-500 hover:bg-red-50 p-1.5 rounded-lg transition-colors cursor-pointer"
                              title="Hapus Item"
                            >
                              <span className="material-symbols-outlined text-[16px]">delete</span>
                            </button>
                          </td>
                        </tr>

                        {/* DRILL DOWN EXPANDABLE DETAIL VIEW */}
                        {expandedProposalId === proposal.id && (
                          <tr className="bg-slate-50/40">
                            <td colSpan={8} className="px-6 py-4 border-b border-[#f1f5f9]">
                              <div className="bg-white p-5 rounded-2xl border border-[#154be2]/10 shadow-[0_4px_24px_rgba(21,75,226,0.02)] grid grid-cols-1 md:grid-cols-3 gap-6 animate-in slide-in-from-top-1.5 duration-250">
                                <div>
                                  <h4 className="text-[10px] font-bold text-[#154be2] uppercase tracking-wider mb-2.5 flex items-center gap-1.5 border-b border-slate-100 pb-1.5">
                                    <span className="material-symbols-outlined text-[15px]">info</span>
                                    Informasi Proyek
                                  </h4>
                                  <div className="space-y-1.5 text-xs">
                                    <div className="flex justify-between">
                                      <span className="text-[#8E94B7] font-semibold">Nomor Proyek:</span>
                                      <span className="font-mono font-bold text-[#154be2]">{proposal.projectNo}</span>
                                    </div>
                                    <div className="flex justify-between">
                                      <span className="text-[#8E94B7] font-semibold">Jenis Kegiatan:</span>
                                      <span className="font-bold text-[#181a2c]">{proposal.activity}</span>
                                    </div>
                                    <div className="flex justify-between">
                                      <span className="text-[#8E94B7] font-semibold">Diajukan Oleh (BS):</span>
                                      <span className="font-bold text-[#181a2c]">{proposal.bs}</span>
                                    </div>
                                  </div>
                                </div>

                                <div>
                                  <h4 className="text-[10px] font-bold text-[#154be2] uppercase tracking-wider mb-2.5 flex items-center gap-1.5 border-b border-slate-100 pb-1.5">
                                    <span className="material-symbols-outlined text-[15px]">map</span>
                                    Wilayah & Waktu
                                  </h4>
                                  <div className="space-y-1.5 text-xs">

                                    <div className="flex justify-between">
                                      <span className="text-[#8E94B7] font-semibold">Bulan Pelaksanaan:</span>
                                      <span className="font-bold text-[#181a2c]">{proposal.bulan}</span>
                                    </div>
                                    <div className="flex justify-between">
                                      <span className="text-[#8E94B7] font-semibold">Tanggal Diajukan:</span>
                                      <span className="font-bold text-slate-500">{proposal.createdAt}</span>
                                    </div>
                                  </div>
                                </div>

                                <div>
                                  <h4 className="text-[10px] font-bold text-[#154be2] uppercase tracking-wider mb-2.5 flex items-center gap-1.5 border-b border-slate-100 pb-1.5">
                                    <span className="material-symbols-outlined text-[15px]">target</span>
                                    Target & Fokus
                                  </h4>
                                  <div className="space-y-1.5 text-xs">
                                    <div className="flex justify-between">
                                      <span className="text-[#8E94B7] font-semibold">Farmer Reach (Target):</span>
                                      <span className="font-bold text-[#181a2c]">{proposal.farmerReach ? Number(proposal.farmerReach).toLocaleString() : "0"} orang</span>
                                    </div>
                                    <div>
                                      <span className="text-[#8E94B7] font-semibold block mb-1">Varietas Fokus:</span>
                                      <div className="flex flex-wrap gap-1 mt-0.5">
                                        {proposal.hybrids.split(", ").map((h: string) => (
                                          <span key={h} className="text-[9px] font-bold text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-100/50">
                                            {h}
                                          </span>
                                        ))}
                                      </div>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={8} className="px-5 py-12 text-center text-[#8E94B7] font-semibold text-xs">
                        Belum ada project yang disubmit. Silakan propose kegiatan di atas.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}


      <LogoutConfirmModal
        isOpen={isLogoutModalOpen}
        onClose={() => setIsLogoutModalOpen(false)}
        onConfirm={onLogout}
      />
    </div>
  );
};

// Login Screen Component
const LoginScreen = ({ onLogin }) => {
  const [name, setName] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleLoginSubmit = async () => {
    setError("");
    setLoading(true);

    try {
      const result = await onLogin(name, password);
      if (!result.success) {
        setError(result.error || "Username atau password salah.");
      }
    } catch {
      setError("Terjadi kesalahan jaringan.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen supports-[min-height:100dvh]:min-h-[100dvh] font-sans flex items-center justify-center p-6 bg-gradient-to-br from-[#F2E7FE] via-[#fbf8ff] to-[#edecff]">
      <div className="glass-panel w-full max-w-sm rounded-[24px] p-8 shadow-[0_20px_40px_rgba(24,26,44,0.06)] border border-white/60 animate-in fade-in zoom-in-95 duration-500">
        <div className="size-24 bg-gradient-to-br from-primary to-cyan-400 rounded-[28px] mx-auto mb-5 shadow-[0_12px_32px_rgba(21,75,226,0.3)] flex items-center justify-center text-white p-3">
          <AdvantaLogo className="w-[64px] h-[64px] text-white" />
        </div>
        <h1 className="text-xl font-bold text-center text-[#181a2c] tracking-tight mb-1 uppercase">
          RADAR DG
        </h1>
        <p className="text-[10px] text-center text-[#8E94B7] font-semibold uppercase tracking-widest mb-6 leading-relaxed">
          REKAN RADAR DG DAN ANALISA REPORT
        </p>

        <div className="space-y-4">
          {error && (
            <div className="bg-red-50 text-red-500 font-semibold text-xs px-4 py-2.5 rounded-full border border-red-100 flex items-center gap-2 animate-in fade-in duration-300">
              <span className="material-symbols-outlined text-sm shrink-0">
                error
              </span>
              <span className="truncate">{error}</span>
            </div>
          )}

          {/* Manual Form */}
          <div className="space-y-4">
            <div>
              <label className="text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider ml-4 mb-2 block">
                username
              </label>
              <input
                type="text"
                value={name}
                onChange={(e) => setName(e.target.value)}
                disabled={loading}
                className="w-full h-14 bg-white/80 border-0 rounded-full px-6 font-semibold text-sm text-[#181a2c] outline-none focus:bg-white transition-all shadow-[0_4px_18px_rgba(21,75,226,0.08)] focus:shadow-[0_8px_28px_rgba(21,75,226,0.18)] mb-4 disabled:opacity-50"
                placeholder="Enter username..."
              />
              <label className="text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider ml-4 mb-2 block">
                Password
              </label>
              <input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                disabled={loading}
                className="w-full h-14 bg-white/80 border-0 rounded-full px-6 font-semibold text-sm text-[#181a2c] outline-none focus:bg-white transition-all shadow-[0_4px_18px_rgba(21,75,226,0.08)] focus:shadow-[0_8px_28px_rgba(21,75,226,0.18)] disabled:opacity-50"
                placeholder="Enter password..."
              />
            </div>

            <button
              onClick={handleLoginSubmit}
              disabled={
                name.trim() === "" || password.trim() === "" || loading
              }
              className={`w-full h-14 rounded-full font-semibold text-xs uppercase tracking-wider transition-all active:scale-[0.98] cursor-pointer flex items-center justify-center gap-2 ${
                name.trim() === "" || password.trim() === "" || loading
                  ? "bg-[#e0e0fa] text-[#8E94B7] cursor-not-allowed shadow-none"
                  : "bg-gradient-to-r from-primary to-cyan-400 text-white hover:opacity-95 shadow-[0_12px_28px_rgba(21,75,226,0.35)]"
              }`}
            >
              {loading ? (
                <>
                  <svg
                    className="animate-spin h-5 w-5 text-white"
                    fill="none"
                    viewBox="0 0 24 24"
                  >
                    <circle
                      className="opacity-25"
                      cx="12"
                      cy="12"
                      r="10"
                      stroke="currentColor"
                      strokeWidth="4"
                    />
                    <path
                      className="opacity-75"
                      fill="currentColor"
                      d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                    />
                  </svg>
                  <span>Logging In...</span>
                </>
              ) : (
                <span>Enter Now</span>
              )}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

const OFFLINE_EMPLOYEES = [];

const OFFLINE_KIOSKS = [];

const OFFLINE_WORKING_DATA = [];

const OFFLINE_DR_SALES = [];

const INDO_MONTHS = [
  "Januari",
  "Februari",
  "Maret",
  "April",
  "Mei",
  "Juni",
  "Juli",
  "Agustus",
  "September",
  "Oktober",
  "November",
  "Desember",
];

const getActivityFullName = (name: string): string => {
  switch (name) {
    case "FFD": return "Farmer Field Day";
    case "FM": return "Farmer Meeting";
    case "ODP": return "One Day Promo";
    case "SFT": return "Special Field Trip";
    case "BFFD": return "Big Farmer Field Day";
    case "BFM": return "Big Farmer Meeting";
    case "BFT": return "Big Field Trip";
    case "BC": return "Branding Crop";
    case "CRV": return "Caravan";
    case "EXP": return "Expo";
    case "PT": return "Pasar Tani";
    default: return name;
  }
};

const formatTooltipValue = (num: number, metricType: string) => {
  if (num === 0) return "0";
  const isNegative = num < 0;
  const absNum = Math.abs(num);
  let result = "";
  if (metricType === "nominal") {
    if (absNum >= 1000000000) result = (absNum / 1000000000).toFixed(1) + " Miliar";
    else if (absNum >= 1000000) result = (absNum / 1000000).toFixed(0) + " Juta";
    else result = Math.round(absNum).toLocaleString("id-ID");
  } else if (metricType === "reach") {
    result = Math.round(absNum).toLocaleString("id-ID");
  } else {
    result = Math.round(absNum).toLocaleString("id-ID") + " x";
  }
  return isNegative ? `-${result}` : result;
};

const formatChartLabelValue = (val: any, metricType: string) => {
  if (val === undefined || val === null || isNaN(Number(val))) return "";
  const num = Number(val);
  if (num === 0) return "0";
  if (metricType === "nominal") {
    if (num >= 1000000000) return (num / 1000000000).toFixed(1) + " M";
    if (num >= 1000000) return (num / 1000000).toFixed(0) + " Jt";
    return Math.round(num).toLocaleString("id-ID");
  } else if (metricType === "reach") {
    if (num >= 1000000) return (num / 1000000).toFixed(1) + "M";
    if (num >= 1000) return (num / 1000).toFixed(1) + "K";
    return Math.round(num).toLocaleString("id-ID");
  }
  return Math.round(num).toLocaleString("id-ID");
};

const CustomBudgetLabel = (props: any) => {
  const { x, y, width, value, metricType } = props;
  if (value === undefined || value === null || value === 0) return null;
  const formattedVal = formatChartLabelValue(value, metricType || "activity");
  const cx = x + width / 2;
  return (
    <g>
      {/* Background white outline/halo for high contrast */}
      <text
        x={cx}
        y={y - 22}
        textAnchor="middle"
        fill="none"
        stroke="#ffffff"
        strokeWidth={4.5}
        strokeLinejoin="round"
        fontSize={11}
        fontWeight={800}
        fontFamily="sans-serif"
      >
        {formattedVal}
      </text>
      <text
        x={cx}
        y={y - 22}
        textAnchor="middle"
        fill="#154be2"
        fontSize={11}
        fontWeight={800}
        fontFamily="sans-serif"
      >
        {formattedVal}
      </text>
    </g>
  );
};

const CustomActualLabel = (props: any) => {
  const { x, y, width, value, metricType } = props;
  if (value === undefined || value === null || value === 0) return null;
  const formattedVal = formatChartLabelValue(value, metricType || "activity");
  const cx = x + width / 2;
  return (
    <g>
      {/* Background white outline/halo for high contrast */}
      <text
        x={cx}
        y={y - 8}
        textAnchor="middle"
        fill="none"
        stroke="#ffffff"
        strokeWidth={4.5}
        strokeLinejoin="round"
        fontSize={11}
        fontWeight={800}
        fontFamily="sans-serif"
      >
        {formattedVal}
      </text>
      <text
        x={cx}
        y={y - 8}
        textAnchor="middle"
        fill="#0a90a6"
        fontSize={11}
        fontWeight={800}
        fontFamily="sans-serif"
      >
        {formattedVal}
      </text>
    </g>
  );
};

const CustomChartTooltip = (props: any) => {
  const { active, payload, label, metricType, dismissedLabel } = props;
  if (dismissedLabel === label) return null;
  if (!active || !payload || !payload.length) return null;

  const data = payload[0].payload;
  if (!data) return null;

  const budget = metricType === "nominal"
    ? (data.budgetNominal || 0)
    : metricType === "reach"
    ? (data.budgetReach || 0)
    : (data.budgetActivity || 0);

  const actual = metricType === "nominal"
    ? (data.actualNominal || 0)
    : metricType === "reach"
    ? (data.actualReach || 0)
    : (data.actualActivity || 0);

  const gap = actual - budget;
  const pct = budget > 0 ? (actual / budget) * 100 : 0;

  const displayName = getActivityFullName(data.name || label || "");

  const formattedBudget = formatTooltipValue(budget, metricType);
  const formattedActual = formatTooltipValue(actual, metricType);
  const gapSign = gap > 0 ? "+" : "";
  const formattedGap = gap === 0 ? "0" : `${gapSign}${formatTooltipValue(gap, metricType)}`;
  const formattedPct = `${Math.round(pct)}%`;

  return (
    <div className="bg-white/95 backdrop-blur-md p-4 rounded-2xl border border-slate-100 shadow-[0_12px_32px_rgba(21,75,226,0.12)] flex flex-col gap-2 min-w-[200px]">
      <div className="font-bold text-[13px] text-slate-800 border-b border-slate-50 pb-1.5 mb-1">
        {displayName}
      </div>
      <div className="flex items-center justify-between gap-4 text-xs">
        <span className="text-slate-400 font-medium">Budget:</span>
        <span className="text-slate-700 font-semibold">{formattedBudget}</span>
      </div>
      <div className="flex items-center justify-between gap-4 text-xs">
        <span className="text-slate-400 font-medium">Actual:</span>
        <span className="text-slate-700 font-semibold">{formattedActual}</span>
      </div>
      <div className="flex items-center justify-between gap-4 text-xs border-t border-slate-50/50 pt-1.5">
        <span className="text-slate-400 font-medium">Gap:</span>
        <span className={`font-bold ${gap >= 0 ? "text-[#0a90a6]" : "text-[#df1b1b]"}`}>
          {formattedGap}
        </span>
      </div>
      <div className="flex items-center justify-between gap-4 text-xs">
        <span className="text-slate-400 font-medium">Persentase:</span>
        <span className="text-slate-700 font-bold bg-slate-50 px-1.5 py-0.5 rounded text-[10px]">
          {formattedPct}
        </span>
      </div>
    </div>
  );
};

const CustomXAxisTick = (props: any) => {
  const { x, y, payload } = props;
  const value = payload.value || "";
  const words = typeof value === "string" ? value.split(" ") : [String(value)];
  
  // Find item in chartData to calculate GAP and Percentage
  const item = props.chartData
    ? (props.chartData.find((d: any) => d.name === value) || props.chartData[props.index])
    : null;

  const actual = item 
    ? (props.metricType === "nominal" 
      ? item.actualNominal 
      : props.metricType === "reach" 
      ? item.actualReach 
      : item.actualActivity) 
    : 0;
  const budget = item 
    ? (props.metricType === "nominal" 
      ? item.budgetNominal 
      : props.metricType === "reach" 
      ? item.budgetReach 
      : item.budgetActivity) 
    : 0;

  const gap = actual - budget;
  const pct = budget > 0 ? (actual / budget) * 100 : 0;

  const formatValue = (num: number, mType: string) => {
    if (num === 0) return "0";
    const absNum = Math.abs(num);
    let result = "";
    if (mType === "nominal") {
      if (absNum >= 1000000000) result = (absNum / 1000000000).toFixed(1) + " M";
      else if (absNum >= 1000000) result = (absNum / 1000000).toFixed(0) + " Jt";
      else result = Math.round(absNum).toLocaleString("id-ID");
    } else if (mType === "reach") {
      if (absNum >= 1000000) result = (absNum / 1000000).toFixed(1) + "M";
      else if (absNum >= 1000) result = (absNum / 1000).toFixed(1) + "K";
      else result = Math.round(absNum).toLocaleString("id-ID");
    } else {
      result = Math.round(absNum).toLocaleString("id-ID");
    }
    return num < 0 ? `-${result}` : result;
  };

  const gapSign = gap > 0 ? "+" : "";
  const formattedGap = gap === 0 ? "0" : `${gapSign}${formatValue(gap, props.metricType)}`;
  const formattedPct = `${Math.round(pct)}%`;
  
  const gapText = item ? `${formattedGap} (${formattedPct})` : "";
  
  // Display name words (limit to 2 lines to save space)
  const nameLines = words.slice(0, 2);
  if (words.length > 2) {
    nameLines[1] = nameLines[1] + "...";
  }

  return (
    <g transform={`translate(${x},${y})`}>
      <text
        x={0}
        y={0}
        textAnchor="middle"
        style={{ fontFamily: "sans-serif" }}
      >
        {item && (
          <tspan
            x={0}
            dy={8}
            fill={gap >= 0 ? "#0a90a6" : "#df1b1b"}
            style={{ fontSize: "11px", fontWeight: 900 }}
          >
            {gapText}
          </tspan>
        )}
        {nameLines.map((word: string, index: number) => (
          <tspan 
            x={0} 
            dy={index === 0 ? (item ? 13 : 8) : 10} 
            key={`word-${index}`} 
            fill="#4e5572" 
            style={{ fontSize: "8.5px", fontWeight: 700 }}
          >
            {word}
          </tspan>
        ))}
      </text>
    </g>
  );
};

const CustomYAxisTick = (props: any) => {
  const { x, y, payload } = props;
  const value = payload.value || "";
  const words = typeof value === "string" ? value.split(" ") : [String(value)];
  
  if (words.length <= 1) {
    return (
      <g transform={`translate(${x},${y})`}>
        <text
          x={-6}
          y={0}
          dy={3}
          textAnchor="end"
          fill="#8E94B7"
          style={{ fontSize: "8.5px", fontWeight: 700, fontFamily: "sans-serif" }}
        >
          {value}
        </text>
      </g>
    );
  }

  const mid = Math.ceil(words.length / 2);
  const line1 = words.slice(0, mid).join(" ");
  const line2 = words.slice(mid).join(" ");

  return (
    <g transform={`translate(${x},${y})`}>
      <text
        x={-6}
        y={0}
        textAnchor="end"
        fill="#8E94B7"
        style={{ fontSize: "8px", fontWeight: 700, fontFamily: "sans-serif" }}
      >
        <tspan x={-6} dy={-2}>{line1}</tspan>
        <tspan x={-6} dy={9}>{line2}</tspan>
      </text>
    </g>
  );
};

export default function App() {
  const [userData, setUserData] = useState(() => {
    try {
      const saved = localStorage.getItem("radar_user_session");
      if (saved) return JSON.parse(saved);
    } catch (e) {
      console.error("Failed to load user session", e);
    }
    // Automatically log in as RADAR DG Admin to bypass login screen
    return {
      name: "RADAR DG Admin",
      email: "admin",
      user: "admin",
      position: "Business Analyst",
      province: "Head Office",
      area: "Head Office",
      level: "5",
      group: "All"
    };
  });


  const saveUserSession = (data: any) => {
    try {
      if (data) {
        localStorage.setItem("radar_user_session", JSON.stringify(data));
        localStorage.removeItem("radar_logged_out");

        // Set activeTab to overview on login if they are a Business Analyst or Admin
        const cleanPos = cleanForMatch(data.position || "");
        const cleanName = cleanForMatch(data.name || "");
        const isAdmin = data.level && String(data.level).toLowerCase().trim() === "admin";
        const isBA = cleanPos === "businessanalyst" || cleanName === "adityawiratama" || cleanName === "aditya" || isAdmin;
        if (isBA) {
          setActiveTab("overview");
        } else {
          setActiveTab("overview");
        }
      } else {
        localStorage.removeItem("radar_user_session");
        localStorage.setItem("radar_logged_out", "true");
        setActiveTab("overview");
      }
    } catch (e) {
      console.error("Failed to save user session", e);
    }
    setUserData(data);
  };

  const handleLogout = () => {
    // Keep user logged in as Admin even if they click logout
    const adminUser = {
      name: "RADAR DG Admin",
      email: "admin",
      user: "admin",
      position: "Business Analyst",
      province: "Head Office",
      area: "Head Office",
      level: "5",
      group: "All"
    };
    saveUserSession(adminUser);
    try {
      localStorage.removeItem('appAccessRules');
    } catch (e) {
      console.error('Failed to clear appAccessRules on logout', e);
    }
    setAccessRules({
      "Business Analyst": { home: true, partner: true, stock: true, pog: true, overview: true, temp: true, access: true },
      "Vegetables Sales Manager": { home: true, partner: true, stock: true, pog: true, overview: true, temp: true, access: false },
      "Commercial Lead": { home: true, partner: true, stock: true, pog: true, overview: true, temp: true, access: false },
      "Country Head": { home: true, partner: true, stock: true, pog: true, overview: true, temp: true, access: true },
      "Area Sales Manager": { home: true, partner: true, stock: true, pog: true, overview: false, temp: false, access: false },
      "Sales Agronomist": { home: true, partner: true, stock: true, pog: true, overview: false, temp: false, access: false },
      "Business Solution": { home: true, partner: true, stock: true, pog: true, overview: false, temp: false, access: false },
    });
  };
  const [activeTab, setActiveTab] = useState(() => {
    try {
      const savedSession = localStorage.getItem("radar_user_session");
      if (savedSession) {
        const parsed = JSON.parse(savedSession);
        const name = parsed.name || "";
        const cleanName = cleanForMatch(name);
        const isAditya = cleanName === "adityawiratama" || cleanName === "aditya";
        const position = parsed.position || "";
        const cleanPos = cleanForMatch(position);

        const isAdmin = parsed.level && String(parsed.level).toLowerCase().trim() === "admin";
        const isBA = cleanPos === "businessanalyst" || isAditya || isAdmin;
        if (isBA) {
          const savedRules = localStorage.getItem('appAccessRules');
          if (savedRules) {
            const rules = JSON.parse(savedRules);
            const matchedKey = Object.keys(rules).find(k => cleanForMatch(k) === "businessanalyst" || cleanForMatch(k) === "aditya" || cleanForMatch(k) === "adityawiratama" || cleanForMatch(k) === "admin");
            if (matchedKey && rules[matchedKey]?.overview === false) {
              return "partner";
            }
          }
          return "overview";
        }
      }
    } catch (e) {
      console.warn("Failed to determine initial tab:", e);
    }
    return "overview";
  });
  const [isMenuVisible, setIsMenuVisible] = useState(true);
  const [isAutoLoggingIn, setIsAutoLoggingIn] = useState(true);

  const greeting = useMemo(() => {
    const hour = new Date().getHours();
    if (hour >= 4 && hour < 11) {
      return {
        text: "Selamat Pagi",
        imageUrl:
          "https://lh3.googleusercontent.com/d/1AzKb-75MaU9hppqSdy2rS93t0tAPGkGi",
        color: "text-amber-300",
      };
    }
    if (hour >= 11 && hour < 15) {
      return {
        text: "Selamat Siang",
        imageUrl:
          "https://lh3.googleusercontent.com/d/1ZpNkT7R57FppIpyPuTt2w9QtJdIwwuRp",
        color: "text-yellow-300",
      };
    }
    if (hour >= 15 && hour < 19) {
      return {
        text: "Selamat Sore",
        imageUrl:
          "https://lh3.googleusercontent.com/d/12RsJXxDrH7aIAph0AJubB3i4w0gmkxcL",
        color: "text-orange-400",
      };
    }
    return {
      text: "Selamat Malam",
      imageUrl:
        "https://lh3.googleusercontent.com/d/1wzqPdQ5jvw7fOF2X76kM56l9l-4mUcLx",
      color: "text-indigo-200",
    };
  }, []);

  const isBusinessAnalyst = useMemo(() => {
    if (!userData) return false;
    const isBA = (userData.position &&
        cleanForMatch(userData.position) === "businessanalyst") ||
      cleanForMatch(userData.name || "") === "adityawiratama" ||
      cleanForMatch(userData.name || "") === "aditya";
    const isAdmin = userData.level && String(userData.level).toLowerCase().trim() === "admin";
    return isBA || isAdmin;
  }, [userData]);

  // Filter states for Executive Overview Tab
  const [overviewMetricFilter, setOverviewMetricFilter] = useState<
    | "POG"
    | "Opening"
    | "sales"
    | "material"
    | "movement"
    | "idle"
    | "total_stock"
    | "activity"
    | "nominal"
    | "reach"
  >("activity");

  // Filter states for the lower part ("yang dibawah")
  const [filterBelowMonth, setFilterBelowMonth] = useState<string[]>([]);
  const [filterBelowChannel, setFilterBelowChannel] = useState<string>("All");
  const [filterBelowMaterial, setFilterBelowMaterial] = useState<string>("All");
  const [filterBelowTeam, setFilterBelowTeam] = useState<string>("All");
  const [filterBelowArea, setFilterBelowArea] = useState<string>("All");
  const [filterBelowCrop, setFilterBelowCrop] = useState<string>("All");
  const [filterBelowType, setFilterBelowType] = useState<string>("All");
  const [isSidebarExpanded, setIsSidebarExpanded] = useState(true);

  const isAditya =
    userData &&
    (cleanForMatch(userData.name) === "adityawiratama" ||
      cleanForMatch(userData.name) === "aditya");

  const userLevel = useMemo(() => {
    if (!userData) return 0;
    if (
      userData.level !== undefined &&
      userData.level !== null &&
      String(userData.level).trim() !== ""
    ) {
      const parsed = parseLevelStr(userData.level);
      if (!isNaN(parsed)) return parsed;
    }
    const rank = getPositionRank(userData.position);
    if (rank === 1) return 5;
    if (rank === 2) return 4;
    if (rank === 3) return 3;
    if (rank === 4) return 2;
    if (rank === 5) return 1;
    return 0;
  }, [userData]);

  const userPosition = useMemo(() => {
    return userData ? normalizePosition(userData.position) : "";
  }, [userData]);

  const [accessRules, setAccessRules] = useState<Record<string, Record<string, boolean>>>(() => {
    try {
      const saved = localStorage.getItem('appAccessRules');
      if (saved) return JSON.parse(saved);
    } catch (e) {
      console.error('Failed to load access rules from localStorage', e);
    }
    return {
      "Business Analyst": { home: true, partner: true, stock: true, pog: true, overview: true, temp: true, access: true },
      "Vegetables Sales Manager": { home: true, partner: true, stock: true, pog: true, overview: true, temp: true, access: false },
      "Commercial Lead": { home: true, partner: true, stock: true, pog: true, overview: true, temp: true, access: false },
      "Country Head": { home: true, partner: true, stock: true, pog: true, overview: true, temp: true, access: true },
      "Area Sales Manager": { home: true, partner: true, stock: true, pog: true, overview: false, temp: false, access: false },
      "Sales Agronomist": { home: true, partner: true, stock: true, pog: true, overview: false, temp: false, access: false },
      "Business Solution": { home: true, partner: true, stock: true, pog: true, overview: false, temp: false, access: false },
    };
  });

  const userAccess = useMemo(() => {
    return {
      home: false,
      partner: true,
      stock: true,
      pog: true,
      overview: true,
      temp: true,
      access: true,
      propose: true,
    };
  }, []);

  const showHomeTab = userData ? !!userAccess.home : false;
  const showPartnerTab = false;
  const showStockTab = userData ? !!userAccess.stock : false;
  const showPogTab = userData ? !!userAccess.pog : false;
  const showOverviewTab = userData ? !!userAccess.overview : false;
  const showTempTab = userData ? !!userAccess.temp : false;
  const showAccessTab = userData ? (!!userAccess.access || isAditya) : false;
  const showProposeTab = userData ? !!userAccess.propose : false;

  // Eager redirection on render to avoid layout flashing and guarantee seamless first login redirection
  if (userData) {
    const isCurrentTabForbidden = 
      (activeTab === "home" && !showHomeTab) ||
      (activeTab === "partner" && !showPartnerTab) ||
      (activeTab === "summary" && !showStockTab) ||
      (activeTab === "pog" && !showPogTab) ||
      (activeTab === "temp" && !showTempTab) ||
      (activeTab === "overview" && !showOverviewTab) ||
      (activeTab === "access" && !showAccessTab) ||
      (activeTab === "propose" && !showProposeTab);

    if (isCurrentTabForbidden) {
      let targetTab = "";
      if (showHomeTab) targetTab = "home";
      else if (showOverviewTab) targetTab = "overview";
      else if (showPartnerTab) targetTab = "partner";
      else if (showStockTab) targetTab = "summary";
      else if (showPogTab) targetTab = "pog";
      else if (showTempTab) targetTab = "temp";
      else if (showAccessTab) targetTab = "access";
      else if (showProposeTab) targetTab = "propose";

      if (targetTab && targetTab !== activeTab) {
        setActiveTab(targetTab);
      }
    }
  }

  // Safety check to redirect from unauthorized or disabled tabs
  useEffect(() => {
    if (userData) {
      if (
        (!showHomeTab && activeTab === "home") ||
        (!showPartnerTab && activeTab === "partner") ||
        (!showStockTab && activeTab === "summary") ||
        (!showPogTab && activeTab === "pog") ||
        (!showTempTab && activeTab === "temp") ||
        (!showOverviewTab && activeTab === "overview") ||
        (!showAccessTab && activeTab === "access") ||
        (!showProposeTab && activeTab === "propose")
      ) {
        // Find first available tab
        if (showHomeTab) setActiveTab("home");
        else if (showOverviewTab) setActiveTab("overview");
        else if (showPartnerTab) setActiveTab("partner");
        else if (showStockTab) setActiveTab("summary");
        else if (showPogTab) setActiveTab("pog");
        else if (showTempTab) setActiveTab("temp");
        else if (showAccessTab) setActiveTab("access");
        else if (showProposeTab) setActiveTab("propose");
      }
    }
  }, [userData, activeTab, showHomeTab, showPartnerTab, showStockTab, showPogTab, showTempTab, showOverviewTab, showAccessTab, showProposeTab]);

  // Load Google Material Symbols for icons
  useEffect(() => {
    const link = document.createElement("link");
    link.href =
      "https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@24,400,0,0";
    link.rel = "stylesheet";
    document.head.appendChild(link);
  }, []);

  const handleLogin = async (name, password) => {
    // Attempt to real login using Apps Script endpoint
    try {
      const resp = await customFetch(
        `${SCRIPT_URL}?action=getUserProfile&user=${encodeURIComponent(name)}`,
      );
      const res = await resp.json();

      if (res.status === "success" && res.data) {
        const data = { ...res.data };
        const dbPassword = String(data.password || "").trim();
        const inputPassword = String(password || "").trim();

        if (inputPassword !== dbPassword) {
          return { success: false, error: "Password salah." };
        }

        const isAditya =
          cleanForMatch(name) === "adityawiratama" ||
          cleanForMatch(name) === "aditya" ||
          cleanForMatch(data.name || "") === "adityawiratama" ||
          cleanForMatch(data.name || "") === "aditya" ||
          cleanForMatch(data.user || "") === "aditya" ||
          cleanForMatch(data.user || "") === "adityawiratama";
        if (isAditya) {
          data.position = "Business Analyst";
        } else {
          data.position = normalizePosition(data.position);
        }

        // Fetch and set actual access rules on login to prevent flashing of unauthorized tabs
        try {
          const accessResp = await customFetch(`${SCRIPT_URL}?action=getAccessRules`);
          const accessRes = await accessResp.json();
          if (accessRes.status === "success" && accessRes.data && Object.keys(accessRes.data).length > 0) {
            setAccessRules(accessRes.data);
            try {
              localStorage.setItem('appAccessRules', JSON.stringify(accessRes.data));
            } catch (e) {
              console.error('Failed to save appAccessRules on login', e);
            }
          }
        } catch (err) {
          console.warn("Failed to pre-fetch access rules during login:", err);
        }

        saveUserSession(data);
        return { success: true };
      } else {
        return {
          success: false,
          error: res.message || "Username tidak ditemukan.",
        };
      }
    } catch (e) {
      console.warn("Login call error:", e);
      return { success: false, error: "Terjadi kesalahan jaringan." };
    }
  };

  useEffect(() => {
    const runAutoLogin = async () => {
      try {
        const savedSession = localStorage.getItem("radar_user_session");

        if (savedSession) {
          const parsed = JSON.parse(savedSession);
          setUserData(parsed);
        }
      } catch (err) {
        console.warn("Auto login error:", err);
      } finally {
        setIsAutoLoggingIn(false);
      }
    };
    runAutoLogin();
  }, []);

  if (isAutoLoggingIn && !userData) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen supports-[min-height:100dvh]:min-h-[100dvh] gap-6 bg-gradient-to-br from-[#F2E7FE] via-[#fbf8ff] to-[#edecff] text-[#181a2c]">
        <div className="relative">
          <div className="size-20 border-4 border-[#edecff] rounded-full"></div>
          <div className="size-20 border-4 border-primary border-t-transparent rounded-full animate-spin absolute inset-0"></div>
          <div className="absolute inset-0 flex items-center justify-center">
            <span className="material-symbols-outlined text-[32px] text-primary animate-pulse">
              sync
            </span>
          </div>
        </div>
        <div className="text-center animate-pulse">
          <h3 className="text-[#181a2c] font-bold text-sm mb-1">
            Membuka Workspace...
          </h3>
          <p className="text-[#8E94B7] text-[9px] font-semibold uppercase tracking-widest">
            Memeriksa Sesi Pengguna
          </p>
        </div>
      </div>
    );
  }

  if (!userData) {
    return <LoginScreen onLogin={handleLogin} />;
  }

  return (
    <div
      className={`min-h-screen supports-[min-height:100dvh]:min-h-[100dvh] bg-[#fbf8ff] font-sans selection:bg-[#edecff] text-[#181a2c] flex flex-col md:flex-row md:items-start w-full mx-auto relative transition-all duration-300`}
    >
      {/* Desktop Sidebar (Visible on md and larger) */}
      <div
        className={`hidden md:flex flex-col ${isSidebarExpanded ? "w-20 lg:w-64" : "w-20"} bg-gradient-to-b from-[#154be2]/[0.1] via-[#154be2]/[0.05] to-white/80 backdrop-blur-2xl border border-[#154be2]/15 h-[calc(100vh-2rem)] sticky top-4 z-50 transition-all duration-300 rounded-3xl ml-4 shadow-[0_8px_32px_rgba(21,75,226,0.15)] overflow-hidden`}
      >

        <div className="flex flex-col gap-2.5 h-[200px] bg-gradient-to-br from-[#154be2] to-cyan-500">

          {/* Sidebar Greeting Card */}
          <div 
            onClick={() => setIsSidebarExpanded(!isSidebarExpanded)}
            className="p-4 rounded-3xl shrink-0 text-white flex flex-col items-center text-center overflow-visible cursor-pointer"
          >
            {isSidebarExpanded ? (
              <div className="hidden lg:flex flex-col rounded-2xl select-none relative overflow-visible items-center text-center">
                <div className="flex items-center gap-3 mb-4">
                  <AdvantaLogo className="size-6 text-white" />
                  <UserIcon className="size-8 rounded-full object-cover" />
                  <span className="font-extrabold text-[12px] tracking-wide text-white uppercase">RADAR DG</span>
                </div>
                <div className="relative z-10 flex flex-col items-center gap-2">
                  {greeting.imageUrl && (
                    <img
                      src={greeting.imageUrl}
                      referrerPolicy="no-referrer"
                      className="size-16 object-contain shrink-0 pointer-events-none select-none animate-rotate-sway"
                      alt=""
                    />
                  )}
                  <div className="min-w-0">
                    <p className="text-[9.5px] text-white/80 font-extrabold uppercase tracking-widest leading-none mb-1">
                      {greeting.text},
                    </p>
                    <h4 className="text-sm font-black text-white leading-tight truncate">
                      {userData.name}
                    </h4>
                  </div>
                </div>
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center py-2 gap-2">
                <UserIcon className="size-8 rounded-full object-cover" />
              </div>
            )}
          </div>
        </div>

        <div className="flex flex-col gap-2.5 px-2.5 lg:px-4 pb-6 bg-white/30 border-t border-[#154be2]/10 pt-6">

          {/* Primary CTA: Input Activity - Keluar dari Group */}
          {showHomeTab ? (
            <button
              onClick={() => setActiveTab("home")}
              className={`flex items-center justify-center lg:justify-start gap-3 h-13 rounded-xl transition-all duration-200 cursor-pointer ${
                activeTab === "home"
                  ? "bg-gradient-to-r from-[#154be2] to-cyan-500 text-white font-extrabold shadow-[0_6px_20px_rgba(21,75,226,0.3)] ring-1 ring-[#154be2]/20 scale-[1.02]"
                  : "bg-gradient-to-r from-[#154be2]/10 to-cyan-400/10 hover:from-[#154be2]/15 hover:to-cyan-400/15 text-[#154be2] border border-[#154be2]/20 font-bold"
              }`}
            >
              <span className="material-symbols-outlined ml-0 lg:ml-4">
                edit_note
              </span>
              <span className={`font-extrabold text-xs hidden ${isSidebarExpanded ? "lg:block" : ""}`}>
                Input Activity
              </span>
            </button>
          ) : showOverviewTab ? (
            <button
              onClick={() => setActiveTab("overview")}
              className={`flex items-center justify-center lg:justify-start gap-3 h-13 rounded-xl transition-all duration-200 cursor-pointer ${
                activeTab === "overview"
                  ? "bg-gradient-to-r from-[#154be2] to-cyan-500 text-white font-extrabold shadow-[0_6px_20px_rgba(21,75,226,0.3)] ring-1 ring-[#154be2]/20 scale-[1.02]"
                  : "bg-gradient-to-r from-[#154be2]/10 to-cyan-400/10 hover:from-[#154be2]/15 hover:to-cyan-400/15 text-[#154be2] border border-[#154be2]/20 font-bold"
              }`}
            >
              <span className="material-symbols-outlined ml-0 lg:ml-4">
                analytics
              </span>
              <span className={`font-extrabold text-xs hidden ${isSidebarExpanded ? "lg:block" : ""}`}>
                Executive Overview
              </span>
            </button>
          ) : null}

          {/* Divider */}
          <div className="h-[1px] bg-[#154be2]/10 my-1 lg:mx-2" />

          {/* Main Navigation Group */}
          {showOverviewTab && showHomeTab && (
            <button
              onClick={() => setActiveTab("overview")}
              className={`flex items-center justify-center lg:justify-start gap-3 h-13 rounded-xl transition-all ${activeTab === "overview" ? "bg-[#154be2]/15 text-[#154be2] shadow-[0_4px_12px_rgba(21,75,226,0.12)] ring-1 ring-[#154be2]/15 font-bold" : "text-[#8E94B7] hover:bg-white/40 hover:text-[#181a2c]"}`}
            >
              <span
                className={`material-symbols-outlined ml-0 lg:ml-4 ${activeTab === "overview" ? "font-normal" : ""}`}
              >
                analytics
              </span>
              <span className={`font-semibold text-xs hidden ${isSidebarExpanded ? "lg:block" : ""}`}>
                Executive Overview
              </span>
            </button>
          )}

          {showPartnerTab && (
            <button
              onClick={() => setActiveTab("partner")}
              className={`flex items-center justify-center lg:justify-start gap-3 h-13 rounded-xl transition-all ${activeTab === "partner" ? "bg-[#154be2]/15 text-[#154be2] shadow-[0_4px_12px_rgba(21,75,226,0.12)] ring-1 ring-[#154be2]/15 font-bold" : "text-[#8E94B7] hover:bg-white/40 hover:text-[#181a2c]"}`}
            >
              <span
                className={`material-symbols-outlined ml-0 lg:ml-4 ${activeTab === "partner" ? "font-normal" : ""}`}
              >
                handshake
              </span>
              <span className={`font-semibold text-xs hidden ${isSidebarExpanded ? "lg:block" : ""}`}>
                Data Partner
              </span>
            </button>
          )}

          {showStockTab && (
            <button
              onClick={() => setActiveTab("summary")}
              className={`flex items-center justify-center lg:justify-start gap-3 h-13 rounded-xl transition-all ${activeTab === "summary" ? "bg-[#154be2]/15 text-[#154be2] shadow-[0_4px_12px_rgba(21,75,226,0.12)] ring-1 ring-[#154be2]/15 font-bold" : "text-[#8E94B7] hover:bg-white/40 hover:text-[#181a2c]"}`}
            >
              <span
                className={`material-symbols-outlined ml-0 lg:ml-4 ${activeTab === "summary" ? "font-normal" : ""}`}
              >
                donut_large
              </span>
              <span className={`font-semibold text-xs hidden ${isSidebarExpanded ? "lg:block" : ""}`}>
                Stock Summary
              </span>
            </button>
          )}

          {showPogTab && (
            <button
              onClick={() => setActiveTab("pog")}
              className={`flex items-center justify-center lg:justify-start gap-3 h-13 rounded-xl transition-all ${activeTab === "pog" ? "bg-[#154be2]/15 text-[#154be2] shadow-[0_4px_12px_rgba(21,75,226,0.12)] ring-1 ring-[#154be2]/15 font-bold" : "text-[#8E94B7] hover:bg-white/40 hover:text-[#181a2c]"}`}
            >
              <span
                className={`material-symbols-outlined ml-0 lg:ml-4 ${activeTab === "pog" ? "font-normal" : ""}`}
              >
                trending_up
              </span>
              <span className={`font-semibold text-xs hidden ${isSidebarExpanded ? "lg:block" : ""}`}>
                POG Tracking
              </span>
            </button>
          )}



          {showAccessTab && (
            <button
              onClick={() => setActiveTab("access")}
              className={`flex items-center justify-center lg:justify-start gap-3 h-13 rounded-xl transition-all ${activeTab === "access" ? "bg-[#154be2]/15 text-[#154be2] shadow-[0_4px_12px_rgba(21,75,226,0.12)] ring-1 ring-[#154be2]/15 font-bold" : "text-[#8E94B7] hover:bg-white/40 hover:text-[#181a2c]"}`}
            >
              <span
                className={`material-symbols-outlined ml-0 lg:ml-4 ${activeTab === "access" ? "font-normal" : ""}`}
              >
                admin_panel_settings
              </span>
              <span className={`font-semibold text-xs hidden ${isSidebarExpanded ? "lg:block" : ""}`}>
                Access Menu
              </span>
            </button>
          )}

          {showProposeTab && (
            <button
              onClick={() => setActiveTab("propose")}
              className={`flex items-center justify-center lg:justify-start gap-3 h-13 rounded-xl transition-all ${activeTab === "propose" ? "bg-[#154be2]/15 text-[#154be2] shadow-[0_4px_12px_rgba(21,75,226,0.12)] ring-1 ring-[#154be2]/15 font-bold" : "text-[#8E94B7] hover:bg-white/40 hover:text-[#181a2c]"}`}
            >
              <span
                className={`material-symbols-outlined ml-0 lg:ml-4 ${activeTab === "propose" ? "font-normal" : ""}`}
              >
                rate_review
              </span>
              <span className={`font-semibold text-xs hidden ${isSidebarExpanded ? "lg:block" : ""}`}>
                Propose Activity
              </span>
            </button>
          )}
        </div>

      </div>

      <div className="flex-1 w-full min-w-0 max-w-2xl lg:max-w-4xl xl:max-w-5xl mx-auto shadow-2xl md:shadow-none bg-[#fbf8ff] relative overflow-hidden pb-24 md:pb-8">
        <Dashboard
          userData={userData}
          activeTab={activeTab}
          onLogout={handleLogout}
          onUserSwitch={handleLogin}
          setUserData={setUserData}
          setActiveTab={setActiveTab}
          accessRules={accessRules}
          setAccessRules={setAccessRules}
          overviewMetricFilter={overviewMetricFilter}
          setOverviewMetricFilter={setOverviewMetricFilter}
          filterBelowMonth={filterBelowMonth}
          setFilterBelowMonth={setFilterBelowMonth}
          filterBelowChannel={filterBelowChannel}
          setFilterBelowChannel={setFilterBelowChannel}
          filterBelowMaterial={filterBelowMaterial}
          setFilterBelowMaterial={setFilterBelowMaterial}
          filterBelowTeam={filterBelowTeam}
          setFilterBelowTeam={setFilterBelowTeam}
          filterBelowArea={filterBelowArea}
          setFilterBelowArea={setFilterBelowArea}
          filterBelowCrop={filterBelowCrop}
          setFilterBelowCrop={setFilterBelowCrop}
          filterBelowType={filterBelowType}
          setFilterBelowType={setFilterBelowType}
        />
      </div>

      {/* Bottom Navigation (Mobile Only) */}
      <div
        className={`md:hidden fixed bottom-3 left-4 right-4 max-w-sm mx-auto flex items-end justify-between gap-3 z-50 transition-all duration-300 ${isMenuVisible ? "scale-100 opacity-100 pointer-events-auto" : "scale-95 opacity-0 pointer-events-none"}`}
      >
        {/* Main Tab Group */}
        <div className="flex-1 flex flex-col bg-gradient-to-b from-[#154be2]/[0.12] via-white/80 to-white/95 backdrop-blur-xl py-1 shadow-[0_12px_36px_rgba(21,75,226,0.12)] rounded-[24px] border border-white/60 relative">
          {/* Hide Button Pill */}
          <button
            onClick={() => setIsMenuVisible(false)}
            className="absolute -top-3 right-6 bg-white/90 backdrop-blur-sm text-[#8E94B7] hover:text-[#181a2c] size-6 rounded-full border border-blue-100/55 shadow-sm transition-all hover:scale-105 active:scale-95 flex items-center justify-center z-50 cursor-pointer"
            title="Sembunyikan Menu"
          >
            <span className="material-symbols-outlined text-xs font-bold">
              keyboard_arrow_down
            </span>
          </button>

          <div className="flex flex-row items-center justify-around h-13 px-2">
            {showOverviewTab && showHomeTab && (
              <button
                onClick={() => setActiveTab("overview")}
                className={`flex flex-col items-center justify-center h-11 px-2.5 rounded-xl transition-all duration-200 select-none ${
                  activeTab === "overview"
                    ? "bg-[#154be2]/12 text-[#154be2] font-extrabold"
                    : "text-[#8E94B7] hover:text-[#181a2c]"
                }`}
              >
                <span
                  className={`material-symbols-outlined text-[18px] leading-tight ${activeTab === "overview" ? "font-semibold" : ""}`}
                >
                  analytics
                </span>
                <span className="text-[7.5px] font-bold uppercase tracking-wider leading-none mt-0.5">
                  Overview
                </span>
              </button>
            )}

            {showPartnerTab && (
              <button
                onClick={() => setActiveTab("partner")}
                className={`flex flex-col items-center justify-center h-11 px-2.5 rounded-xl transition-all duration-200 select-none ${
                  activeTab === "partner"
                    ? "bg-[#154be2]/12 text-[#154be2] font-extrabold"
                    : "text-[#8E94B7] hover:text-[#181a2c]"
                }`}
              >
                <span
                  className={`material-symbols-outlined text-[18px] leading-tight ${activeTab === "partner" ? "font-semibold" : ""}`}
                >
                  handshake
                </span>
                <span className="text-[7.5px] font-bold uppercase tracking-wider leading-none mt-0.5">
                  Partner
                </span>
              </button>
            )}

            {showStockTab && (
              <button
                onClick={() => setActiveTab("summary")}
                className={`flex flex-col items-center justify-center h-11 px-2.5 rounded-xl transition-all duration-200 select-none ${
                  activeTab === "summary"
                    ? "bg-[#154be2]/12 text-[#154be2] font-extrabold"
                    : "text-[#8E94B7] hover:text-[#181a2c]"
                }`}
              >
                <span
                  className={`material-symbols-outlined text-[18px] leading-tight ${activeTab === "summary" ? "font-semibold" : ""}`}
                >
                  donut_large
                </span>
                <span className="text-[7.5px] font-bold uppercase tracking-wider leading-none mt-0.5">
                  Stock
                </span>
              </button>
            )}

            {showPogTab && (
              <button
                onClick={() => setActiveTab("pog")}
                className={`flex flex-col items-center justify-center h-11 px-2.5 rounded-xl transition-all duration-200 select-none ${
                  activeTab === "pog"
                    ? "bg-[#154be2]/12 text-[#154be2] font-extrabold"
                    : "text-[#8E94B7] hover:text-[#181a2c]"
                }`}
              >
                <span
                  className={`material-symbols-outlined text-[18px] leading-tight ${activeTab === "pog" ? "font-semibold" : ""}`}
                >
                  trending_up
                </span>
                <span className="text-[7.5px] font-bold uppercase tracking-wider leading-none mt-0.5">
                  POG
                </span>
              </button>
            )}

            {showTempTab && (
              <button
                onClick={() => setActiveTab("temp")}
                className={`flex flex-col items-center justify-center h-11 px-2.5 rounded-xl transition-all duration-200 select-none ${
                  activeTab === "temp"
                    ? "bg-[#154be2]/12 text-[#154be2] font-extrabold"
                    : "text-[#8E94B7] hover:text-[#181a2c]"
                }`}
              >
                <span
                  className={`material-symbols-outlined text-[19px] leading-tight ${activeTab === "temp" ? "font-semibold" : ""}`}
                >
                  assignment
                </span>
                <span className="text-[7.5px] font-bold uppercase tracking-wider leading-none mt-0.5">
                  Temp
                </span>
              </button>
            )}



            {showAccessTab && (
              <button
                onClick={() => setActiveTab("access")}
                className={`flex flex-col items-center justify-center h-11 px-2.5 rounded-xl transition-all duration-200 select-none ${
                  activeTab === "access"
                    ? "bg-[#154be2]/12 text-[#154be2] font-extrabold"
                    : "text-[#8E94B7] hover:text-[#181a2c]"
                }`}
              >
                <span
                  className={`material-symbols-outlined text-[19px] leading-tight ${activeTab === "access" ? "font-semibold" : ""}`}
                >
                  admin_panel_settings
                </span>
                <span className="text-[7.5px] font-bold uppercase tracking-wider leading-none mt-0.5">
                  Access
                </span>
              </button>
            )}

            {showProposeTab && (
              <button
                onClick={() => setActiveTab("propose")}
                className={`flex flex-col items-center justify-center h-11 px-2.5 rounded-xl transition-all duration-200 select-none ${
                  activeTab === "propose"
                    ? "bg-[#154be2]/12 text-[#154be2] font-extrabold"
                    : "text-[#8E94B7] hover:text-[#181a2c]"
                }`}
              >
                <span
                  className={`material-symbols-outlined text-[19px] leading-tight ${activeTab === "propose" ? "font-semibold" : ""}`}
                >
                  rate_review
                </span>
                <span className="text-[7.5px] font-bold uppercase tracking-wider leading-none mt-0.5">
                  Propose
                </span>
              </button>
            )}
          </div>
        </div>

        {/* Separate Input Button - Floating Action Button Style */}
        {showHomeTab ? (
          <button
            onClick={() => setActiveTab("home")}
            className={`flex items-center justify-center shrink-0 size-[56px] rounded-full transition-all duration-200 select-none shadow-[0_8px_20px_rgba(21,75,226,0.35)] active:scale-95 cursor-pointer text-white bg-gradient-to-tr from-[#154be2] to-cyan-500`}
          >
            <span className="material-symbols-outlined text-[26px]">
              edit_note
            </span>
          </button>
        ) : showOverviewTab ? (
          <button
            onClick={() => setActiveTab("overview")}
            className={`flex items-center justify-center shrink-0 size-[56px] rounded-full transition-all duration-200 select-none shadow-[0_8px_20px_rgba(21,75,226,0.35)] active:scale-95 cursor-pointer text-white bg-gradient-to-tr from-[#154be2] to-cyan-500`}
          >
            <span className="material-symbols-outlined text-[26px]">
              analytics
            </span>
          </button>
        ) : null}
      </div>

      {/* Show Menu Trigger */}
      {!isMenuVisible && (
        <button
          onClick={() => setIsMenuVisible(true)}
          className="md:hidden fixed bottom-4 right-4 bg-gradient-to-r from-primary to-cyan-400 text-white size-9 rounded-full shadow-[0_6px_20px_rgba(21,75,226,0.25)] flex items-center justify-center hover:opacity-95 active:scale-[0.98] transition-all z-50 animate-in fade-in slide-in-from-bottom-4 duration-300 cursor-pointer"
          title="Tampilkan Menu"
        >
          <span className="material-symbols-outlined text-base">
            keyboard_arrow_up
          </span>
        </button>
      )}
    </div>
  );
}
