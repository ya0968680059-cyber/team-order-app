import { useEffect, useMemo, useRef, useState, type ChangeEvent } from 'react';
import * as XLSX from 'xlsx';

type MenuVariant = { id: string; name: string; price: number };
type MenuAddOn = { id: string; name: string; price: number };
type MenuCustom = { id: string; name: string };

type MenuItem = {
  id: string;
  name: string; // 主餐名稱
  variants: MenuVariant[]; // 規格（單選）
  addOns: MenuAddOn[]; // 加價配料（多選）
  customs: MenuCustom[]; // 免費客製（多選）
};

type MenuLibraryEntry = { id: string; name: string; items: MenuItem[]; createdAt: string };

type OrderRecord = {
  id: string;
  buyer: string;
  mainName: string;
  variant: MenuVariant;
  addOns: MenuAddOn[];
  customs: MenuCustom[];
  qty: number;
  note?: string;
  total: number; // 計算後的總價 (variant + addons) * qty
  createdAt: string;
  paid: boolean;
  changeAmount: number; // 找零金額（而非 boolean）
  own: boolean;
};

const STORAGE_KEY = 'TEAM_ORDER_SHARED_DATA_V2';
const MY_IDS_KEY = 'TEAM_ORDER_MY_IDS_V2';
const MENU_LIBRARY_KEY = 'TEAM_ORDER_MENU_LIBRARY_V2';
const CURRENT_MENU_KEY = 'TEAM_ORDER_CURRENT_MENU_V2';
const DEFAULT_API = 'https://keyvalue.imanyou.com/api/';
const API_KEY = 'team-order-app-v2';
const MANAGER_PASSWORD = '#309';
const MANAGER_PASSWORD_KEY = 'TEAM_ORDER_MANAGER_PASSWORD_V2';

const initialMenu: MenuItem[] = [
  {
    id: 'm-1',
    name: '招牌便當',
    variants: [
      { id: 'v-1', name: '一般', price: 120 },
      { id: 'v-2', name: '大盛', price: 150 },
    ],
    addOns: [{ id: 'a-1', name: '爆炒羊', price: 20 }],
    customs: [{ id: 'c-1', name: '換蛋全熟' }],
  },
  {
    id: 'm-2',
    name: '雞腿飯',
    variants: [{ id: 'v-3', name: '一般', price: 150 }],
    addOns: [],
    customs: [],
  },
];

const getApiUrl = () => `${DEFAULT_API}${API_KEY}`;

async function loadRemoteOrders() {
  try {
    const res = await fetch(getApiUrl(), { method: 'GET' });
    if (!res.ok) return [];
    const data = await res.json();
    if (Array.isArray(data)) return data;
    if (data && Array.isArray((data as any).orders)) return (data as any).orders;
    return [];
  } catch {
    return [];
  }
}

async function saveRemoteOrders(orders: OrderRecord[]) {
  try {
    await fetch(getApiUrl(), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(orders),
    });
  } catch {
    // ignore network errors and keep local fallback
  }
}

function App() {
  const [menuLibrary, setMenuLibrary] = useState<MenuLibraryEntry[]>(() => {
    try {
      const raw = localStorage.getItem(MENU_LIBRARY_KEY);
      if (raw) {
        const parsed = JSON.parse(raw) as MenuLibraryEntry[];
        if (Array.isArray(parsed) && parsed.length) return parsed;
      }
    } catch {
      // ignore
    }
    return [{ id: 'default', name: '預設菜單', items: initialMenu, createdAt: new Date().toLocaleString('zh-TW') }];
  });

  const [selectedLibraryId, setSelectedLibraryId] = useState<string>(() => menuLibrary[0]?.id ?? 'default');
  const [menu, setMenu] = useState<MenuItem[]>(() => menuLibrary[0]?.items ?? initialMenu);

  const [role, setRole] = useState<'colleague' | 'manager'>('colleague');
  const [managerPassword, setManagerPassword] = useState<string>(() => {
    try {
      return localStorage.getItem(MANAGER_PASSWORD_KEY) || MANAGER_PASSWORD;
    } catch {
      return MANAGER_PASSWORD;
    }
  });

  const [orders, setOrders] = useState<OrderRecord[]>(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) return JSON.parse(raw) as OrderRecord[];
    } catch {}
    return [];
  });

  const [myIds, setMyIds] = useState<string[]>(() => {
    try {
      const raw = localStorage.getItem(MY_IDS_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch {
      return [];
    }
  });

  // colleague UI state
  const [buyer, setBuyer] = useState('');
  const [menuSearch, setMenuSearch] = useState('');
  const [menuDropdownOpen, setMenuDropdownOpen] = useState(false);
  const [selectedMainId, setSelectedMainId] = useState<string | null>(null);
  const [selectedVariantId, setSelectedVariantId] = useState<string | null>(null);
  const [selectedAddOnIds, setSelectedAddOnIds] = useState<string[]>([]);
  const [selectedCustomIds, setSelectedCustomIds] = useState<string[]>([]);
  const [quantity, setQuantity] = useState<number>(1);
  const [itemNote, setItemNote] = useState('');

  const [uploading, setUploading] = useState(false);
  const [syncState, setSyncState] = useState('未同步');

  // 管理員驗證時間戳（15分鐘免密碼進入）
  const [lastAdminVerifiedTime, setLastAdminVerifiedTime] = useState<number>(() => {
    try {
      const stored = localStorage.getItem('TEAM_ORDER_LAST_ADMIN_TIME_V2');
      return stored ? parseInt(stored, 10) : 0;
    } catch {
      return 0;
    }
  });

  // 外送便條紙信息
  const [deliveryOrderer, setDeliveryOrderer] = useState('經濟部產業園區管理局 / 徐梓恩');
  const [deliveryAddress, setDeliveryAddress] = useState('高雄市楠梓區加昌路600號');
  const [deliveryPhone, setDeliveryPhone] = useState('(07)361-1212 #309');
  const [deliveryTime, setDeliveryTime] = useState('中午 11 點半');
  const [deliveryNote, setDeliveryNote] = useState('送餐時可以直接臨停汽車坡道，進來後直接跟櫃台說分機 309 訂餐，謝謝您！');

  const dropdownRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    loadRemoteOrders().then((remoteOrders) => {
      if (remoteOrders.length) {
        setOrders(remoteOrders as OrderRecord[]);
        localStorage.setItem(STORAGE_KEY, JSON.stringify(remoteOrders));
        setSyncState('已同步遠端資料');
      } else {
        setSyncState('未找到遠端資料');
      }
    });
  }, []);

  useEffect(() => {
    localStorage.setItem(MY_IDS_KEY, JSON.stringify(myIds));
  }, [myIds]);

  useEffect(() => {
    localStorage.setItem(MENU_LIBRARY_KEY, JSON.stringify(menuLibrary));
  }, [menuLibrary]);

  useEffect(() => {
    localStorage.setItem(MANAGER_PASSWORD_KEY, managerPassword);
  }, [managerPassword]);

  useEffect(() => {
    localStorage.setItem(CURRENT_MENU_KEY, JSON.stringify(menu));
  }, [menu]);

  useEffect(() => {
    localStorage.setItem('TEAM_ORDER_LAST_ADMIN_TIME_V2', String(lastAdminVerifiedTime));
  }, [lastAdminVerifiedTime]);

  useEffect(() => {
    const handleOutsideClick = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setMenuDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', handleOutsideClick);
    return () => document.removeEventListener('mousedown', handleOutsideClick);
  }, []);

  const switchToManager = () => {
    const now = Date.now();
    const timeDiffMs = now - lastAdminVerifiedTime;
    const fifteenMinutesMs = 15 * 60 * 1000;

    // 如果在15分鐘內已驗證過，直接進入管理頁面
    if (timeDiffMs < fifteenMinutesMs) {
      setRole('manager');
      setLastAdminVerifiedTime(now); // 更新驗證時間
      return;
    }

    // 超過15分鐘或未驗證過，彈出密碼輸入框
    const password = window.prompt('輸入主揪密碼以進入管理頁面：');
    if (password === managerPassword) {
      setRole('manager');
      setLastAdminVerifiedTime(now); // 更新驗證時間
    } else {
      alert('密碼錯誤');
    }
  };

  const normalizeMenuLibrary = (entries: MenuLibraryEntry[]) =>
    entries.length ? entries : [{ id: 'default', name: '預設菜單', items: initialMenu, createdAt: new Date().toLocaleString('zh-TW') }];

  const switchMenuLibrary = (entryId: string) => {
    const entry = menuLibrary.find((e) => e.id === entryId);
    if (entry) {
      setSelectedLibraryId(entry.id);
      setMenu(entry.items);
      setSyncState(`已切換到 ${entry.name}`);
    }
  };

  const clearOrders = () => {
    if (window.confirm('確認清除本次訂單並開啟新團購？')) {
      setOrders([]);
      setMyIds([]);
      setSyncState('已清除訂單並開啟新團購');
      saveRemoteOrders([]);
    }
  };

  const changeManagerPassword = () => {
    const next = window.prompt('請輸入新的管理密碼：', '');
    if (next && next.trim()) {
      setManagerPassword(next.trim());
      alert('管理密碼已更新');
    }
  };

  const downloadMenuTemplate = () => {
    const csv = [
      ['主餐', '規格', '規格價格', '加價配料', '加價價格', '免費客製'].join(','),
      ['美國雪花牛', '小盛', '180', '爆炒羊', '20', '換蛋全熟'].join(','),
      ['美國雪花牛', '大盛', '220', '', '', '換蛋半熟'].join(','),
      ['奶茶', '一般', '45', '', '', '少冰'].join(','),
    ].join('\n');
    // Prepend UTF-8 BOM so Excel opens CSV with correct encoding for Chinese
    const blob = new Blob(['\uFEFF', csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'menu-structured-template.csv';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const deleteMenuEntry = (entryId: string) => {
    const nextLibrary = normalizeMenuLibrary(menuLibrary.filter((e) => e.id !== entryId));
    setMenuLibrary(nextLibrary);
    if (selectedLibraryId === entryId) {
      setSelectedLibraryId(nextLibrary[0].id);
      setMenu(nextLibrary[0].items);
    }
    setSyncState('已刪除菜單');
  };

  const importMenu = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet) as Array<Record<string, unknown>>;

      const map = new Map<string, MenuItem>();
      rows.forEach((row) => {
        const main = String(row['主餐'] ?? row['品名'] ?? row['Name'] ?? row['name'] ?? '').trim();
        if (!main) return;
        if (!map.has(main)) {
          map.set(main, { id: `m-${Date.now()}-${map.size}`, name: main, variants: [], addOns: [], customs: [] });
        }
        const entry = map.get(main)!;

        const variantName = String(row['規格'] ?? row['Variant'] ?? row['variant'] ?? '').trim();
        const variantPriceRaw = row['規格價格'] ?? row['VariantPrice'] ?? row['價格'] ?? row['Price'] ?? row['price'];
        const variantPrice = Number(variantPriceRaw ?? 0);
        if (variantName) {
          const vid = `v-${main}-${variantName}`;
          if (!entry.variants.find((v) => v.name === variantName)) entry.variants.push({ id: vid, name: variantName, price: Number.isFinite(variantPrice) ? variantPrice : 0 });
        } else if (variantPriceRaw !== undefined && !entry.variants.length) {
          const vid = `v-${main}-一般`;
          if (!entry.variants.find((v) => v.id === vid)) entry.variants.push({ id: vid, name: '一般', price: Number.isFinite(variantPrice) ? variantPrice : 0 });
        }

        const addOnName = String(row['加價配料'] ?? row['AddOn'] ?? row['配料'] ?? '').trim();
        const addOnPrice = Number(row['加價價格'] ?? row['AddOnPrice'] ?? row['配料價格'] ?? 0);
        if (addOnName) {
          const aid = `a-${main}-${addOnName}-${addOnPrice}`;
          if (!entry.addOns.find((a) => a.name === addOnName)) entry.addOns.push({ id: aid, name: addOnName, price: Number.isFinite(addOnPrice) ? addOnPrice : 0 });
        }

        const customName = String(row['免費客製'] ?? row['Custom'] ?? row['客製'] ?? '').trim();
        if (customName) {
          const cid = `c-${main}-${customName}`;
          if (!entry.customs.find((c) => c.name === customName)) entry.customs.push({ id: cid, name: customName });
        }
      });

      const imported = Array.from(map.values());
      if (!imported.length) {
        const rows2 = XLSX.utils.sheet_to_json(sheet) as Array<Record<string, unknown>>;
        rows2.forEach((row, idx) => {
          const name = String(row['品名'] ?? row['name'] ?? row['Name'] ?? '').trim();
          const price = Number(row['價格'] ?? row['price'] ?? 0);
          if (!name) return;
          const id = `m-${Date.now()}-${idx}`;
          imported.push({ id, name, variants: [{ id: `${id}-v-1`, name: '一般', price: Number.isFinite(price) ? price : 0 }], addOns: [], customs: [] });
        });
      }

      if (imported.length) {
        const storeName = file.name.replace(/\.[^.]+$/, '') || '匯入菜單';
        const newEntry: MenuLibraryEntry = { id: `menu-${Date.now()}`, name: storeName, items: imported, createdAt: new Date().toLocaleString('zh-TW') };
        setMenuLibrary((prev) => [newEntry, ...prev.filter((e) => e.id !== newEntry.id)]);
        setMenu(imported);
        setSelectedLibraryId(newEntry.id);
        setSyncState(`已載入 ${storeName}`);
      } else {
        alert('匯入的 Excel/CSV 無法辨識任一主餐');
      }
    } catch (err) {
      console.error(err);
      alert('匯入失敗，請確認檔案格式為 Excel 或 CSV，且欄位包含主餐/規格/加價配料/免費客製');
    }
    setUploading(false);
    event.target.value = '';
  };

  const currentMenuName = useMemo(() => menuLibrary.find((e) => e.id === selectedLibraryId)?.name || '預設菜單', [menuLibrary, selectedLibraryId]);

  const filteredMenu = useMemo(() => {
    const keyword = menuSearch.trim().toLowerCase();
    return keyword ? menu.filter((item) => item.name.toLowerCase().includes(keyword)) : menu;
  }, [menu, menuSearch]);

  const totalAmount = useMemo(() => orders.reduce((s, o) => s + o.total, 0), [orders]);
  const totalQty = useMemo(() => orders.reduce((s, o) => s + o.qty, 0), [orders]);
  const myOrders = useMemo(() => orders.filter((o) => myIds.includes(o.id)), [orders, myIds]);

  const formatOrderLabel = (o: OrderRecord) => {
    const addons = o.addOns.map((a) => a.name).join('、');
    const customs = o.customs.map((c) => c.name).join('、');
    const parts: string[] = [];
    parts.push(`${o.mainName}(${o.variant.name})`);
    if (addons) parts.push(addons);
    if (customs) parts.push(customs);
    return `${parts.join(' + ')} x${o.qty} 份`;
  };

  const submitOrder = () => {
    if (!buyer.trim()) {
      alert('請填寫分機 + 姓名');
      return;
    }
    if (!selectedMainId) {
      alert('請選擇主餐');
      return;
    }
    const main = menu.find((m) => m.id === selectedMainId);
    if (!main) {
      alert('找不到所選主餐，請重新選擇');
      return;
    }
    const variant = main.variants.find((v) => v.id === selectedVariantId) ?? main.variants[0];
    if (!variant) {
      alert('此主餐沒有任何規格，請確認菜單');
      return;
    }
    const addOns = main.addOns.filter((a) => selectedAddOnIds.includes(a.id));
    const customs = main.customs.filter((c) => selectedCustomIds.includes(c.id));
    const base = variant.price + addOns.reduce((s, a) => s + a.price, 0);
    const total = base * quantity;

    const record: OrderRecord = {
      id: `${Date.now()}`,
      buyer: buyer.trim(),
      mainName: main.name,
      variant,
      addOns,
      customs,
      qty: quantity,
      note: itemNote.trim() || undefined,
      total,
      createdAt: new Date().toLocaleString('zh-TW'),
      paid: false,
      changeAmount: 0,
      own: true,
    };

    const next = [record, ...orders];
    setOrders(next);
    setMyIds((prev) => (prev.includes(record.id) ? prev : [record.id, ...prev]));
    setSelectedMainId(null);
    setSelectedVariantId(null);
    setSelectedAddOnIds([]);
    setSelectedCustomIds([]);
    setQuantity(1);
    setItemNote('');
    setMenuSearch('');
    setMenuDropdownOpen(false);
    setSyncState('已提交，正在同步');
    saveRemoteOrders(next);
  };

  const deleteOrder = (id: string) => {
    const next = orders.filter((o) => o.id !== id);
    setOrders(next);
    setMyIds((prev) => prev.filter((x) => x !== id));
    saveRemoteOrders(next);
    setSyncState('已刪除訂單');
  };

  const togglePaid = (id: string) => {
    const next = orders.map((o) => (o.id === id ? { ...o, paid: !o.paid } : o));
    setOrders(next);
    saveRemoteOrders(next);
  };

  const setChangeAmount = (id: string, amount: number) => {
    const next = orders.map((o) => (o.id === id ? { ...o, changeAmount: amount } : o));
    setOrders(next);
    saveRemoteOrders(next);
  };

  const deleteOrderItem = (id: string) => {
    const next = orders.filter((o) => o.id !== id);
    setOrders(next);
    setMyIds((prev) => prev.filter((x) => x !== id));
    saveRemoteOrders(next);
    setSyncState('已刪除訂單項目');
  };

  // 按買方分組訂單
  const ordersByBuyer = useMemo(() => {
    const map = new Map<string, OrderRecord[]>();
    orders.forEach((order) => {
      if (!map.has(order.buyer)) {
        map.set(order.buyer, []);
      }
      map.get(order.buyer)!.push(order);
    });
    return Array.from(map.entries());
  }, [orders]);

  // 商家餐點統計：按「主餐 + 規格 + 加價配料 + 免費客製 + 備註」進行分組
  type ItemKey = string;
  type GroupedItem = {
    key: ItemKey;
    mainName: string;
    variant: MenuVariant;
    addOns: MenuAddOn[];
    customs: MenuCustom[];
    note?: string;
    totalQty: number;
  };

  const groupedItems = useMemo(() => {
    const map = new Map<ItemKey, GroupedItem>();
    orders.forEach((order) => {
      // 創建唯一 key：主餐名稱 + 規格 + 加配料排序後 + 免費客製排序後 + 備註
      const addOnKey = order.addOns
        .map((a) => a.id)
        .sort()
        .join('|');
      const customKey = order.customs
        .map((c) => c.id)
        .sort()
        .join('|');
      const key = `${order.mainName}|${order.variant.id}|${addOnKey}|${customKey}|${order.note || ''}`;

      if (map.has(key)) {
        const item = map.get(key)!;
        item.totalQty += order.qty;
      } else {
        map.set(key, {
          key,
          mainName: order.mainName,
          variant: order.variant,
          addOns: order.addOns,
          customs: order.customs,
          note: order.note,
          totalQty: order.qty,
        });
      }
    });
    // 按主餐名稱進行排序（使用 localeCompare 支援中文排序）
    return Array.from(map.values()).sort((a, b) => a.mainName.localeCompare(b.mainName, 'zh-TW'));
  }, [orders]);

  const formatGroupedItemLabel = (item: GroupedItem) => {
    const addons = item.addOns.map((a) => a.name).join('、');
    const customs = item.customs.map((c) => c.name).join('、');
    const parts: string[] = [];
    parts.push(`${item.mainName}(${item.variant.name})`);
    if (addons) parts.push(addons);
    if (customs) parts.push(customs);
    const label = parts.join(' + ');
    const noteStr = item.note ? ` (備註: ${item.note})` : '';
    return `${label}${noteStr} x ${item.totalQty} 份`;
  };

  // 獨特訂購人數（不同 buyer 的數量）
  const uniqueOrderers = useMemo(() => new Set(orders.map((o) => o.buyer)).size, [orders]);

  const copyDeliveryOrder = async () => {
    const lines: string[] = [];
    lines.push(`訂購人：${deliveryOrderer}`);
    lines.push(`外送地址：${deliveryAddress}`);
    lines.push(`連絡電話：${deliveryPhone}`);
    lines.push(`送餐時間：${deliveryTime}`);
    lines.push(`備註：${deliveryNote}`);
    lines.push('');
    lines.push('餐點統計：');
    groupedItems.forEach((item) => {
      lines.push(formatGroupedItemLabel(item));
    });
    const text = lines.join('\n');
    try {
      await navigator.clipboard.writeText(text);
      alert('已複製訂單資訊');
    } catch {
      alert('複製失敗，請手動複製');
    }
  };

  return (
    <div className="container">
      <div className="card mb">
        <h1 className="title">團體訂餐系統（結構化菜單）</h1>
        <p className="sub">支援主餐 / 規格(單選) / 加價配料(多選) / 免費客製(多選)</p>
        <div className="row mb">
          <span className="tag">同步狀態：{syncState}</span>
          <button
            className="btn secondary"
            onClick={() =>
              loadRemoteOrders().then((remote) => {
                if (remote.length) {
                  setOrders(remote as OrderRecord[]);
                  localStorage.setItem(STORAGE_KEY, JSON.stringify(remote));
                  setSyncState('已重新同步');
                }
              })
            }
          >
            重新同步
          </button>
        </div>

        <div className="row" style={{ gap: 10, flexWrap: 'wrap', marginTop: 16 }}>
          <button 
            className={role === 'colleague' ? 'btn primary' : 'btn secondary'} 
            onClick={() => setRole('colleague')}
            style={{
              flex: 1,
              minWidth: '120px',
              fontWeight: 600,
              transition: 'all 0.3s ease',
              boxShadow: role === 'colleague' ? '0 4px 12px rgba(0, 102, 204, 0.3)' : 'none',
            }}
          >
            🧑‍💻 同事點餐
          </button>
          <button 
            className={role === 'manager' ? 'btn primary' : 'btn secondary'} 
            onClick={switchToManager}
            style={{
              flex: 1,
              minWidth: '120px',
              fontWeight: 600,
              transition: 'all 0.3s ease',
              boxShadow: role === 'manager' ? '0 4px 12px rgba(0, 102, 204, 0.3)' : 'none',
            }}
          >
            🔑 主揪管理
          </button>
        </div>
      </div>

      {role === 'colleague' ? (
        <div className="card">
          <h3>{currentMenuName}</h3>
          <div className="grid" style={{ marginTop: 12 }}>
            <input className="input" placeholder="分機+訂購人（例如：309梓恩）" value={buyer} onChange={(e) => setBuyer(e.target.value)} />

            <div>
              <h4 className="mb">選主餐</h4>
              <div style={{ position: 'relative' }} ref={dropdownRef}>
                <input
                  className="input"
                  placeholder="搜尋或選擇主餐"
                  value={menuSearch}
                  onChange={(e) => {
                    setMenuSearch(e.target.value);
                    setMenuDropdownOpen(true);
                  }}
                  onFocus={() => setMenuDropdownOpen(true)}
                />
                {menuDropdownOpen && (
                  <div style={{ position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 20, background: '#fff', border: '1px solid #ddd', borderRadius: 4, maxHeight: 240, overflowY: 'auto' }}>
                    {filteredMenu.length ? (
                      filteredMenu.map((m) => (
                        <button
                          key={m.id}
                          type="button"
                          className="btn-link"
                          style={{ display: 'block', width: '100%', textAlign: 'left', padding: 10 }}
                          onClick={() => {
                            setSelectedMainId(m.id);
                            setSelectedVariantId(m.variants[0]?.id ?? null);
                            setSelectedAddOnIds([]);
                            setSelectedCustomIds([]);
                            setMenuSearch(m.name);
                            setMenuDropdownOpen(false);
                          }}
                        >
                          {m.name}
                        </button>
                      ))
                    ) : (
                      <div style={{ padding: 12 }}>找不到符合的主餐</div>
                    )}
                  </div>
                )}
              </div>
            </div>

            {selectedMainId ? (
              (() => {
                const main = menu.find((m) => m.id === selectedMainId)!;
                return (
                  <div>
                    <h4 className="mb">規格（單選）</h4>
                    <div>
                      {main.variants.map((v) => (
                        <label key={v.id} style={{ display: 'block', marginBottom: 6 }}>
                          <input type="radio" name="variant" checked={selectedVariantId === v.id} onChange={() => setSelectedVariantId(v.id)} /> {v.name} - ${v.price}
                        </label>
                      ))}
                    </div>

                    {main.addOns.length ? (
                      <>
                        <h4 className="mb">加價配料（多選）</h4>
                        <div>
                          {main.addOns.map((a) => (
                            <label key={a.id} style={{ display: 'block', marginBottom: 6 }}>
                              <input
                                type="checkbox"
                                checked={selectedAddOnIds.includes(a.id)}
                                onChange={() => setSelectedAddOnIds((prev) => (prev.includes(a.id) ? prev.filter((x) => x !== a.id) : [...prev, a.id]))}
                              />
                              {a.name} +${a.price}
                            </label>
                          ))}
                        </div>
                      </>
                    ) : null}

                    {main.customs.length ? (
                      <>
                        <h4 className="mb">免費客製（多選）</h4>
                        <div>
                          {main.customs.map((c) => (
                            <label key={c.id} style={{ display: 'block', marginBottom: 6 }}>
                              <input type="checkbox" checked={selectedCustomIds.includes(c.id)} onChange={() => setSelectedCustomIds((prev) => (prev.includes(c.id) ? prev.filter((x) => x !== c.id) : [...prev, c.id]))} /> {c.name}
                            </label>
                          ))}
                        </div>
                      </>
                    ) : null}

                    <div style={{ marginTop: 8 }}>
                      <label>數量：
                        <select value={String(quantity)} onChange={(e) => setQuantity(Number(e.target.value))} style={{ marginLeft: 8 }}>
                          {Array.from({ length: 10 }, (_, i) => i + 1).map((n) => (
                            <option key={n} value={n}>{n} 份</option>
                          ))}
                        </select>
                      </label>
                    </div>

                    <textarea className="textarea" placeholder="備註（例：小辣、去冰）" value={itemNote} onChange={(e) => setItemNote(e.target.value)} style={{ marginTop: 8 }} />

                    <div style={{ marginTop: 12 }}>
                      <button className="btn success" onClick={submitOrder}>送出訂單</button>
                    </div>
                  </div>
                );
              })()
            ) : null}

            <div className="card" style={{ marginTop: 16 }}>
              <h3 className="mb">我的點餐紀錄</h3>
              {myOrders.length ? (
                <table className="table">
                  <thead>
                    <tr>
                      <th>訂購人</th>
                      <th>內容</th>
                      <th>金額</th>
                      <th>操作</th>
                    </tr>
                  </thead>
                  <tbody>
                    {myOrders.map((order) => (
                      <tr key={order.id}>
                        <td>{order.buyer}</td>
                        <td>{formatOrderLabel(order)}{order.note ? ` (${order.note})` : ''}</td>
                        <td>{order.total}</td>
                        <td>
                          <button className="btn danger" onClick={() => deleteOrder(order.id)}>刪除</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              ) : (
                <div className="small">目前沒有任何你自己送出的訂單。</div>
              )}
            </div>
          </div>
        </div>
      ) : (
        <>
          <div className="card mb" style={{ background: '#f5f7fa', boxShadow: '0 2px 8px rgba(0,0,0,0.08)', borderRadius: '12px', border: '1px solid #e8eef5' }}>
            <h3 className="mb" style={{ color: '#1a202c', fontSize: '1.25rem', fontWeight: 700 }}>主揪控制台</h3>
            <div className="row mb" style={{ gap: 8, flexWrap: 'wrap' }}>
              <button className="btn danger" onClick={clearOrders} style={{ transition: 'all 0.2s', borderRadius: '6px' }}>清除本次訂單，開啟新團購</button>
              <button className="btn secondary" onClick={changeManagerPassword} style={{ transition: 'all 0.2s', borderRadius: '6px' }}>變更管理密碼</button>
            </div>
          </div>

          <div className="card mb" style={{ background: '#fff', boxShadow: '0 2px 8px rgba(0,0,0,0.08)', borderRadius: '12px', border: '1px solid #e8eef5' }}>
            <h3 className="mb" style={{ color: '#1a202c', fontSize: '1.25rem', fontWeight: 700 }}>菜單管理（僅供管理員/主揪操作）</h3>
            <div className="row mb" style={{ alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
              <select className="select" value={selectedLibraryId} onChange={(e) => switchMenuLibrary(e.target.value)} style={{ maxWidth: 320 }}>
                {menuLibrary.map((entry) => (
                  <option key={entry.id} value={entry.id}>{entry.name}</option>
                ))}
              </select>
              <label className="btn secondary">
                匯入 Excel/CSV 菜單
                <input type="file" accept=".xlsx,.xls,.csv" onChange={importMenu} style={{ display: 'none' }} />
              </label>
              <button className="btn secondary" onClick={downloadMenuTemplate}>下載 Excel 菜單範例</button>
              {uploading ? <span className="small">正在匯入…</span> : null}
            </div>
            <div className="small">新版結構：欄位可包含「主餐, 規格, 規格價格, 加價配料, 加價價格, 免費客製」。同一主餐的不同規格會合併為同一主餐節點。</div>

            <div className="card" style={{ marginTop: 12, background: '#f9f9f9', borderRadius: '8px', border: '1px solid #e8e8e8', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
              <h4 className="mb" style={{ color: '#1a202c', fontWeight: 700 }}>已上傳菜單</h4>
              {menuLibrary.length ? (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                  {menuLibrary.map((entry) => (
                    <div
                      key={entry.id}
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'space-between',
                        padding: '8px 12px',
                        borderRadius: '6px',
                        backgroundColor: selectedLibraryId === entry.id ? '#0066cc' : '#e8e8e8',
                        color: selectedLibraryId === entry.id ? '#fff' : '#333',
                        fontSize: '0.95rem',
                        border: 'none',
                      }}
                    >
                      <button
                        type="button"
                        style={{
                          background: 'none',
                          border: 'none',
                          cursor: 'pointer',
                          color: 'inherit',
                          fontSize: 'inherit',
                          padding: 0,
                          marginRight: 8,
                          fontWeight: selectedLibraryId === entry.id ? 'bold' : 'normal',
                        }}
                        onClick={() => switchMenuLibrary(entry.id)}
                      >
                        {entry.name}
                      </button>
                      <button
                        type="button"
                        style={{
                          background: 'none',
                          border: 'none',
                          cursor: 'pointer',
                          color: 'inherit',
                          fontSize: '1.4rem',
                          padding: 0,
                          marginLeft: 4,
                          lineHeight: 1,
                        }}
                        onClick={() => deleteMenuEntry(entry.id)}
                        title="刪除此菜單"
                      >
                        ×
                      </button>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="small">目前沒有任何已上傳的菜單。</div>
              )}
            </div>
          </div>

          <div className="card" style={{ marginTop: 16, background: '#fff', boxShadow: '0 2px 8px rgba(0,0,0,0.08)', borderRadius: '12px', border: '1px solid #e8eef5', padding: '24px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
              <h3 style={{ margin: 0, color: '#1a202c', fontSize: '1.25rem', fontWeight: 700 }}>外送資訊</h3>
              <button className="btn secondary" onClick={copyDeliveryOrder} style={{ transition: 'all 0.2s', borderRadius: '6px', fontWeight: 600 }}>一鍵複製</button>
            </div>

            <div className="grid" style={{ marginBottom: 16 }}>
              <div>
                <label className="mb" style={{ fontSize: '0.95rem', fontWeight: 500, color: '#444' }}>訂購人：</label>
                <input
                  className="input"
                  value={deliveryOrderer}
                  onChange={(e) => setDeliveryOrderer(e.target.value)}
                  style={{ borderRadius: '6px', border: '1px solid #d0d7de', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}
                />
              </div>
              <div>
                <label className="mb" style={{ fontSize: '0.95rem', fontWeight: 500, color: '#444' }}>外送地址：</label>
                <input
                  className="input"
                  value={deliveryAddress}
                  onChange={(e) => setDeliveryAddress(e.target.value)}
                  style={{ borderRadius: '6px', border: '1px solid #d0d7de', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}
                />
              </div>
              <div>
                <label className="mb" style={{ fontSize: '0.95rem', fontWeight: 500, color: '#444' }}>連絡電話：</label>
                <input
                  className="input"
                  value={deliveryPhone}
                  onChange={(e) => setDeliveryPhone(e.target.value)}
                  style={{ borderRadius: '6px', border: '1px solid #d0d7de', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}
                />
              </div>
              <div>
                <label className="mb" style={{ fontSize: '0.95rem', fontWeight: 500, color: '#444' }}>送餐時間：</label>
                <input
                  className="input"
                  value={deliveryTime}
                  onChange={(e) => setDeliveryTime(e.target.value)}
                  style={{ borderRadius: '6px', border: '1px solid #d0d7de', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}
                />
              </div>
              <div>
                <label className="mb" style={{ fontSize: '0.95rem', fontWeight: 500, color: '#444' }}>備註：</label>
                <textarea
                  className="textarea"
                  value={deliveryNote}
                  onChange={(e) => setDeliveryNote(e.target.value)}
                  rows={3}
                  style={{ borderRadius: '6px', border: '1px solid #d0d7de', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}
                />
              </div>
            </div>

            <div className="summary-box mb" style={{ background: '#f0f4f8', borderRadius: '8px', border: '1px solid #cbd5e0', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
              <div style={{ fontSize: '1rem', fontWeight: 600, color: '#2c3e50' }}>訂購人數：{uniqueOrderers}</div>
              <div style={{ fontSize: '1rem', fontWeight: 600, color: '#2c3e50' }}>總份數：{totalQty}</div>
              <div style={{ fontSize: '1rem', fontWeight: 600, color: '#2c3e50' }}>總金額：{totalAmount}</div>
            </div>

            <h4 className="mb" style={{ color: '#1a202c', fontSize: '1.1rem', fontWeight: 700 }}>餐點統計</h4>
            {groupedItems.length ? (
              <div style={{ marginBottom: 16, background: '#f9f9f9', borderRadius: '6px', border: '1px solid #e8e8e8', overflow: 'hidden' }}>
                {groupedItems.map((item, idx) => (
                  <div key={idx} style={{ padding: '12px 16px', borderBottom: idx < groupedItems.length - 1 ? '1px solid #eee' : 'none', fontSize: '0.95rem', backgroundColor: idx % 2 === 0 ? '#fff' : '#fafafa' }}>
                    {formatGroupedItemLabel(item)}
                  </div>
                ))}
              </div>
            ) : (
              <div className="small">目前沒有任何訂單。</div>
            )}
          </div>

          <div className="card" style={{ marginTop: 16, background: '#fff', boxShadow: '0 1px 3px rgba(0,0,0,0.1)', borderRadius: '12px', padding: '24px' }}>
            <h3 className="mb">收款管理</h3>
            {ordersByBuyer.length ? (
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '16px' }}>
                {ordersByBuyer.map(([buyerName, buyerOrders]) => {
                  const buyerTotal = buyerOrders.reduce((sum, o) => sum + o.total, 0);
                  const buyerPaid = buyerOrders.some((o) => o.paid);
                  const firstOrder = buyerOrders[0];
                  return (
                    <div
                      key={buyerName}
                      style={{
                        flex: '1 1 calc(33.333% - 16px)',
                        minWidth: '280px',
                        background: '#f9f9f9',
                        border: '1px solid #e0e0e0',
                        borderRadius: '8px',
                        padding: '16px',
                        boxShadow: '0 2px 4px rgba(0,0,0,0.05)',
                      }}
                    >
                      <h4 style={{ margin: '0 0 12px 0', color: '#333', fontSize: '1.05rem', fontWeight: 'bold' }}>{buyerName}</h4>
                      
                      {/* 餐點列表 */}
                      <div style={{ marginBottom: '12px' }}>
                        {buyerOrders.map((order) => (
                          <div
                            key={order.id}
                            style={{
                              display: 'flex',
                              justifyContent: 'space-between',
                              alignItems: 'flex-start',
                              padding: '8px 0',
                              borderBottom: '1px solid #eee',
                              fontSize: '0.9rem',
                            }}
                          >
                            <div style={{ flex: 1 }}>
                              <span>{formatOrderLabel(order)}</span>
                              {order.note && <span style={{ color: '#666', fontSize: '0.85rem' }}> ({order.note})</span>}
                            </div>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginLeft: '8px' }}>
                              <span style={{ fontWeight: 'bold', minWidth: '50px', textAlign: 'right' }}>${order.total}</span>
                              <button
                                type="button"
                                style={{
                                  background: 'none',
                                  border: 'none',
                                  cursor: 'pointer',
                                  color: '#d32f2f',
                                  fontSize: '1.2rem',
                                  padding: '0 4px',
                                  lineHeight: 1,
                                  transition: 'transform 0.2s',
                                }}
                                onClick={() => deleteOrderItem(order.id)}
                                title="刪除此項目"
                                onMouseEnter={(e) => (e.currentTarget.style.transform = 'scale(1.3)')}
                                onMouseLeave={(e) => (e.currentTarget.style.transform = 'scale(1)')}
                              >
                                ×
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>

                      {/* 總金額 */}
                      <div style={{ fontWeight: 'bold', fontSize: '1.1rem', marginBottom: '12px', paddingTop: '8px' }}>
                        應付總額：<span style={{ textDecoration: buyerPaid ? 'line-through' : 'none' }}>${buyerTotal}</span>
                      </div>

                      {/* 已收款 */}
                      <div style={{ marginBottom: '12px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <input
                          type="checkbox"
                          checked={buyerPaid}
                          onChange={() => {
                            buyerOrders.forEach((o) => togglePaid(o.id));
                          }}
                        />
                        <label style={{ cursor: 'pointer', fontSize: '0.95rem' }}>已收款</label>
                      </div>

                      {/* 找零金額輸入 */}
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <label style={{ fontSize: '0.95rem' }}>找零：</label>
                        <input
                          type="number"
                          min="0"
                          value={firstOrder.changeAmount}
                          onChange={(e) => {
                            const amount = Number(e.target.value) || 0;
                            buyerOrders.forEach((o) => setChangeAmount(o.id, amount));
                          }}
                          style={{
                            flex: 1,
                            padding: '6px 8px',
                            border: '1px solid #ccc',
                            borderRadius: '4px',
                            fontSize: '0.9rem',
                          }}
                          placeholder="0"
                        />
                      </div>
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className="small">目前沒有任何訂單。</div>
            )}
          </div>
        </>
      )}
    </div>
  );
}

export default App;
