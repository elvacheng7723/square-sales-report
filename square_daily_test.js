// 每日Square销售汇总脚本(可用于定时任务自动运行):
// 拉取指定日期的Square订单,按渠道统计薯条营业额 + 按标准化商品名汇总数量,
// 并把结果保存成文件(reports/ 目录下,按日期命名),同时打印到控制台。
//
// 运行方式:
//   SQUARE_ACCESS_TOKEN=你的token node square_daily_test.js
//   也可以指定日期(不指定则默认"西澳今天"): TARGET_DATE=2026-08-09 SQUARE_ACCESS_TOKEN=xxx node square_daily_test.js
// 需要 Node 18+ (自带 fetch)

const fs = require("fs");
const path = require("path");

const ACCESS_TOKEN = process.env.SQUARE_ACCESS_TOKEN;
const LOCATION_ID = "LYNJXKDG9KWXZ";      // 开发者后台 Locations 页面查到的店铺ID
const TIMEZONE_OFFSET = "+08:00";          // 西澳时区(全年不换夏令时,固定UTC+8)
const REPORTS_DIR = path.join(__dirname, "reports");

if (!ACCESS_TOKEN) {
  console.error("请先设置环境变量 SQUARE_ACCESS_TOKEN 再运行本脚本。");
  console.error("例如: SQUARE_ACCESS_TOKEN=xxxx node square_daily_test.js");
  process.exit(1);
}

// 默认取"西澳当前时间"的日期,而不是电脑本地时区的日期
// (定时任务大多跑在UTC服务器上,不能直接用 new Date() 本地日期)
function perthTodayDateStr() {
  const now = new Date();
  // 把当前UTC时间平移8小时,再取 ISO 日期部分,等效于西澳当地日期
  const perthMs = now.getTime() + 8 * 60 * 60 * 1000;
  return new Date(perthMs).toISOString().slice(0, 10);
}

const TARGET_DATE = process.env.TARGET_DATE || perthTodayDateStr();

// 自定义支付方式的名字要和你在Square后台创建时填的完全一致
const CHANNEL_TENDER_NAMES = {
  "Uber Eats": "Uber Eats",
  "DoorDash": "DoorDash",
};

// ============================================================
// 商品映射表 (来自 Items_Corresponding.xlsx)
// ============================================================

// --- 完全忽略的商品(既不计金额也不计数量,也不算未识别) ---
const IGNORED_ITEMS = new Set(
  ["Seasoning Choice"].map((s) => s.trim().toLowerCase())
);

// --- Chips 页: 这些商品只统计 Gross Sales 金额,不统计数量 ---
const CHIPS_ITEMS = new Set(
  [
    "$1 Chips",
    "$2 Chips",
    "$2 Chips (Click to modify size)",
    "$5 Chips",
    "$8 Chips",
    "Chips $2",
    "Chips $8",
  ].map((s) => s.trim().toLowerCase())
);

// --- Items 页: Origin(原始商品名) -> { category, target, ratio } ---
// 原始表里同一 Target 有大小写/空格不同的多种写法,这里统一按 trim+lowerCase 做key
const ITEM_ROWS = [
  ["Drinks", "1.25L Drinks", "1.25L Drinks", 1],
  ["Drinks", "2L Drinks", "2L Drinks", 1],
  ["Drinks", "Bundaburg", "Bundaburg", 1],
  ["Drinks", "Cans", "Cans", 1],
  ["Drinks", "Cans (Voided)", "Cans (Voided)", 1],
  ["Drinks", "ICE TEA 500ml", "ICE TEA 500ml", 1],
  ["Drinks", "Pop-top", "Pop-top", 1],
  ["Drinks", "Water Bottle", "Water Bottle", 1],
  ["Extras", "Pickled Onions(Each)", "Pickled Onions(Each)", 1],
  ["Extras", "Pickled Onions(Jar)", "Pickled Onions(Jar)", 1],
  ["Extras", "Sauce portions", "Sauce portions", 1],
  ["Extras", "Sauce portions - Tartare", "Sauce portions", 1],
  ["Extras", "Sauce portions - Tomato", "Sauce portions", 1],
  ["Extras", "Sauce portions - Sweet Chili Thai", "Sauce portions", 1],
  ["Extras", "Tartare Sauce Jar", "Tartare Sauce Jar", 1],
  ["Extras", "Tomato Sauce Jar", "Tomato Sauce Jar", 1],
  ["Fish", "barramundi", "Barramundi", 1],
  ["Fish", "Barramundi", "Barramundi", 1],
  ["Fish", "goldband Snapper", "Goldband Snapper", 1],
  ["Fish", "Goldband Snapper", "Goldband Snapper", 1],
  ["Fish", "hake Large", "Hake Large", 1],
  ["Fish", "hake Large for Packs", "Hake Large", 1],
  ["Fish", "Hake Large", "Hake Large", 1],
  ["Fish", "Hake Large for Packs", "Hake Large", 1],
  ["Fish", "hake Small", "Hake Small", 1],
  ["Fish", "hake Small for Packs", "Hake Small", 1],
  ["Fish", "Hake Small", "Hake Small", 1],
  ["Fish", "Hake Small for Packs", "Hake Small", 1],
  ["Fish", "shark", "Shark", 1],
  ["Fish", "Shark", "Shark", 1],
  ["Fish", "Silver Whiting", "Silver Whiting", 1],
  ["Fish", "silver Whiting", "Silver Whiting", 1],
  ["Fish", "Spanish Mackerel", "Spanish Mackerel", 1],
  ["Fish", "spanish Mackerel", "Spanish Mackerel", 1],
  ["Old Favourites", "chiko Roll", "Chiko Roll", 1],
  ["Old Favourites", "Chiko Roll", "Chiko Roll", 1],
  ["Old Favourites", "corn Jack", "Corn Jack", 1],
  ["Old Favourites", "Corn Jack", "Corn Jack", 1],
  ["Old Favourites", "Dim Sim", "Dim Sim", 1],
  ["Old Favourites", "dim Sim", "Dim Sim", 1],
  ["Old Favourites", "fish Cake", "Fish Cake", 1],
  ["Old Favourites", "Fish Cake", "Fish Cake", 1],
  ["Old Favourites", "Onion Rings", "Onion Rings", 1],
  ["Old Favourites", "Potato Scallop", "Potato Scallop", 1],
  ["Old Favourites", "potato Scallop", "Potato Scallop", 1],
  ["Old Favourites", "spring Roll", "Spring Roll", 1],
  ["Old Favourites", "Spring Roll", "Spring Roll", 1],
  ["Seafood", "Crab Sticks", "Crab Sticks", 1],
  ["Seafood", "crab Sticks", "Crab Sticks", 1],
  ["Seafood", "fremantle Sardines - 6 per serve", "Fremantle Sardines - 6 per serve", 1],
  ["Seafood", "Fremantle Sardines - 6 per serve", "Fremantle Sardines - 6 per serve", 1],
  ["Seafood", "Mussels", "Mussels", 1],
  ["Seafood", "mussels", "Mussels", 1],
  ["Seafood", "Oysters", "Oysters", 1],
  ["Seafood", "oysters", "Oysters", 1],
  ["Seafood", "Prawns", "Prawns", 1],
  ["Seafood", "prawns", "Prawns", 1],
  ["Seafood", "prawns for packs", "Prawns", 1],
  ["Seafood", "sea Scallops", "Sea Scallops", 1],
  ["Seafood", "Sea Scallops", "Sea Scallops", 1],
  ["Seafood", "sea Scallops for packs", "Sea Scallops", 1],
  ["Seafood", "Squid - 6 per serve", "Squid Ring", 6],
  ["Seafood", "Squid - 6 per serve ", "Squid Ring", 6],
  ["Seafood", "squid - 4 per serve", "Squid Ring", 4],
  ["Something Sweet", "Banana Fritter", "Banana Fritter", 1],
  ["Something Sweet", "Pineapple Fritter", "Pineapple Fritter", 1],
];

const ITEM_MAP = new Map();
for (const [category, origin, target, ratio] of ITEM_ROWS) {
  ITEM_MAP.set(origin.trim().toLowerCase(), {
    category,
    target: target.trim(),
    ratio,
  });
}

// --- Packs 页: 组合商品拆分, 每卖出1份"Seafood Basket"按下面数量拆给对应商品 ---
// 注意: "$2 Chips" 这份套餐内含,Square不会给它单独的销售金额,
// 这里按你说的口径,固定按每份 $2(200分)计入薯条营业额。
const PACKS = {
  "seafood basket": [
    { target: "Hake Small", qty: 2 },
    { target: "Prawns", qty: 2 },
    { target: "Sea Scallops", qty: 1 },
    { target: "Crab Sticks", qty: 1 },
    { target: "Squid Ring", qty: 4 },
    { target: "$2 Chips", qty: 1, isChipsPortion: true, chipsPriceCents: 200 },
  ],
};

// ============================================================
// 拉取订单
// ============================================================

async function fetchAllOrders() {
  const orders = [];
  let cursor = undefined;

  do {
    const body = {
      location_ids: [LOCATION_ID],
      query: {
        filter: {
          // 改用 created_at 筛选,而不是 closed_at:
          // 如果订单是"开台"模式(付款后没有立刻手动关闭),closed_at可能一直是空的,
          // 用 closed_at 筛选会把这些订单整单漏掉。created_at 在订单一创建就有值,更不容易漏单。
          date_time_filter: {
            created_at: {
              start_at: `${TARGET_DATE}T00:00:00${TIMEZONE_OFFSET}`,
              end_at: `${TARGET_DATE}T23:59:59${TIMEZONE_OFFSET}`,
            },
          },
          // 不再只要 COMPLETED,连 OPEN(可能还没手动关台,但已经付过款)也一起拉,
          // 具体是否"已付款"在下面用 tenders 字段再筛一次,避免把没付钱的空单也算进去。
          state_filter: { states: ["COMPLETED", "OPEN"] },
        },
      },
      limit: 100,
      ...(cursor ? { cursor } : {}),
    };

    const res = await fetch("https://connect.squareup.com/v2/orders/search", {
      method: "POST",
      headers: {
        "Square-Version": "2025-01-23",
        "Authorization": `Bearer ${ACCESS_TOKEN}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    const data = await res.json();

    if (!res.ok) {
      console.error("Square API 返回错误:", JSON.stringify(data, null, 2));
      process.exit(1);
    }

    // 只保留真正有付款记录的订单(排除还没结账的空/未付款订单)
    const paidOrders = (data.orders || []).filter(
      (o) => Array.isArray(o.tenders) && o.tenders.length > 0
    );
    orders.push(...paidOrders);
    cursor = data.cursor;
  } while (cursor);

  return orders;
}

// 判断一笔订单属于哪个销售渠道: Uber Eats / DoorDash / Square线上 / 线下
function classifyChannel(order) {
  // 1) 先看有没有专门代表 Uber Eats / DoorDash 的自定义支付方式
  const tenders = order.tenders || [];
  for (const tender of tenders) {
    const noteOrType = tender.type === "OTHER" ? tender.note : tender.type;
    if (!noteOrType) continue;
    const noteOrTypeLower = noteOrType.toLowerCase();
    for (const [channel, tenderName] of Object.entries(CHANNEL_TENDER_NAMES)) {
      if (noteOrTypeLower.includes(tenderName.toLowerCase())) {
        return channel;
      }
    }
  }

  // 2) 再看订单来源,判断是不是通过 Square 自己的网店/线上点餐下的单
  //    (Square Online 生成的订单 source.name 里通常会带 "online" 字样)
  const sourceName = ((order.source && order.source.name) || "").toLowerCase();
  if (sourceName.includes("online")) {
    return "Square线上";
  }

  // 3) 剩下的就是收银机现场打的单
  return "线下";
}

// ============================================================
// 汇总逻辑
// ============================================================

function createState() {
  return {
    chipsSalesCents: 0,
    chipsSalesByChannel: {},
    chipsPortionFromPacksQty: 0, // 来自套餐(如Seafood Basket)的薯条份数,按每份$2计入薯条营业额
    // qtySummary: { target: { category, byChannel: { channel: qty }, total } }
    qtySummary: {},
    unknown: [], // 未在任何映射表中出现的商品
    // channelSummary: { channel: { revenueCents, orderCount } } - 各渠道总销售额(所有商品,不限薯条)
    channelSummary: {},
  };
}

function addQty(state, target, category, channel, qty) {
  if (!state.qtySummary[target]) {
    state.qtySummary[target] = { category, byChannel: {}, total: 0 };
  }
  const entry = state.qtySummary[target];
  entry.byChannel[channel] = (entry.byChannel[channel] || 0) + qty;
  entry.total += qty;
}

function processLineItem(item, channel, state) {
  const rawName = (item.name || "未命名商品").trim();
  const nameKey = rawName.toLowerCase();
  const qty = parseFloat(item.quantity) || 0;

  // 0) 直接忽略的商品
  if (IGNORED_ITEMS.has(nameKey)) {
    return;
  }

  // 1) 薯条: 只统计金额,不统计数量
  if (CHIPS_ITEMS.has(nameKey)) {
    const gross = (item.gross_sales_money && item.gross_sales_money.amount) || 0;
    state.chipsSalesCents += gross;
    state.chipsSalesByChannel[channel] = (state.chipsSalesByChannel[channel] || 0) + gross;
    return;
  }

  // 2) 组合套餐(如 Seafood Basket) 拆分
  const packItems = PACKS[nameKey];
  if (packItems) {
    for (const p of packItems) {
      if (p.isChipsPortion) {
        const portions = p.qty * qty;
        state.chipsPortionFromPacksQty += portions;
        const gross = portions * (p.chipsPriceCents || 0);
        state.chipsSalesCents += gross;
        state.chipsSalesByChannel[channel] = (state.chipsSalesByChannel[channel] || 0) + gross;
        continue;
      }
      addQty(state, p.target, `${rawName} 拆分`, channel, p.qty * qty);
    }
    return;
  }

  // 3) 常规Item映射
  const mapped = ITEM_MAP.get(nameKey);
  if (mapped) {
    addQty(state, mapped.target, mapped.category, channel, qty * mapped.ratio);
    return;
  }

  // 4) 未识别商品,提示需要补充映射表
  state.unknown.push({ name: rawName, qty, channel });
}

function centsToDollarStr(cents) {
  return (cents / 100).toFixed(2);
}

// ============================================================
// 主流程
// ============================================================

async function main() {
  const orders = await fetchAllOrders();

  const state = createState();
  for (const order of orders) {
    const channel = classifyChannel(order);
    const lineItems = order.line_items || [];

    // 统计这笔订单的总营业额(所有商品,不限薯条),按渠道汇总
    if (!state.channelSummary[channel]) {
      state.channelSummary[channel] = { revenueCents: 0, orderCount: 0 };
    }
    state.channelSummary[channel].orderCount += 1;
    for (const item of lineItems) {
      const gross = (item.gross_sales_money && item.gross_sales_money.amount) || 0;
      state.channelSummary[channel].revenueCents += gross;
    }

    for (const item of lineItems) {
      processLineItem(item, channel, state);
    }
  }

  // ============================================================
  // 报表A: 产品维度 - 每个商品卖了多少个(备餐/进货用,不含金额)
  // ============================================================
  const itemLines = [];
  const logItem = (s = "") => itemLines.push(s);

  logItem(`共拉取到 ${orders.length} 笔已完成订单\n`);
  logItem(`=========== ${TARGET_DATE} 商品销量汇总 ===========\n`);

  const byCategory = {};
  for (const [target, entry] of Object.entries(state.qtySummary)) {
    if (!byCategory[entry.category]) byCategory[entry.category] = [];
    byCategory[entry.category].push({ target, ...entry });
  }

  logItem("【商品数量汇总(按渠道拆分)】");
  for (const [category, items] of Object.entries(byCategory)) {
    logItem(`\n-- ${category} --`);
    items
      .sort((a, b) => b.total - a.total)
      .forEach(({ target, byChannel, total }) => {
        const detail = Object.entries(byChannel)
          .map(([ch, q]) => `${ch}: ${q}`)
          .join(", ");
        logItem(`  ${target}  合计 ${total}  (${detail})`);
      });
  }

  if (state.unknown.length > 0) {
    logItem("\n【未识别商品 - 请补充到 Items_Corresponding.xlsx】");
    const grouped = {};
    for (const u of state.unknown) {
      grouped[u.name] = (grouped[u.name] || 0) + u.qty;
    }
    for (const [name, qty] of Object.entries(grouped)) {
      logItem(`  "${name}"  出现数量合计 ${qty}`);
    }
  }

  const itemReportText = itemLines.join("\n");
  console.log(itemReportText);

  // ============================================================
  // 报表B: 渠道维度 - 每个渠道当天卖了多少钱(核算成本/对账Uber&DoorDash用)
  // ============================================================
  const channelLines = [];
  const logChannel = (s = "") => channelLines.push(s);

  logChannel(`共拉取到 ${orders.length} 笔已完成订单\n`);
  logChannel(`=========== ${TARGET_DATE} 渠道营业额汇总 ===========\n`);

  logChannel("【各渠道总营业额】");
  const channelOrder = ["线下", "Square线上", "Uber Eats", "DoorDash"];
  const seenChannels = new Set(channelOrder);
  const allChannelNames = [
    ...channelOrder,
    ...Object.keys(state.channelSummary).filter((c) => !seenChannels.has(c)),
  ];
  let totalRevenueCents = 0;
  for (const ch of allChannelNames) {
    const entry = state.channelSummary[ch];
    if (!entry) continue;
    totalRevenueCents += entry.revenueCents;
    logChannel(`  ${ch}:  $${centsToDollarStr(entry.revenueCents)}  (${entry.orderCount} 笔订单)`);
  }
  logChannel(`  合计:  $${centsToDollarStr(totalRevenueCents)}`);
  logChannel();

  logChannel(`【薯条 Gross Sales 总额】 $${centsToDollarStr(state.chipsSalesCents)}`);
  for (const [channel, cents] of Object.entries(state.chipsSalesByChannel)) {
    logChannel(`  - ${channel}: $${centsToDollarStr(cents)}`);
  }
  if (state.chipsPortionFromPacksQty > 0) {
    logChannel(
      `  其中包含 ${state.chipsPortionFromPacksQty} 份来自 Seafood Basket 等套餐的薯条,` +
        `按每份 $2 计入了上面的金额统计。`
    );
  }

  const channelReportText = channelLines.join("\n");
  console.log(channelReportText);

  // --- 保存结果到文件 ---
  fs.mkdirSync(REPORTS_DIR, { recursive: true });

  // 报表A: 产品维度
  const itemsTxtPath = path.join(REPORTS_DIR, `${TARGET_DATE}-items.txt`);
  fs.writeFileSync(itemsTxtPath, itemReportText, "utf8");

  const itemsJsonPath = path.join(REPORTS_DIR, `${TARGET_DATE}-items.json`);
  fs.writeFileSync(
    itemsJsonPath,
    JSON.stringify(
      {
        date: TARGET_DATE,
        generatedAt: new Date().toISOString(),
        ordersCount: orders.length,
        qtySummary: state.qtySummary,
        unknownItems: state.unknown,
      },
      null,
      2
    ),
    "utf8"
  );

  // 报表B: 渠道维度
  const channelsTxtPath = path.join(REPORTS_DIR, `${TARGET_DATE}-channels.txt`);
  fs.writeFileSync(channelsTxtPath, channelReportText, "utf8");

  const channelsJsonPath = path.join(REPORTS_DIR, `${TARGET_DATE}-channels.json`);
  fs.writeFileSync(
    channelsJsonPath,
    JSON.stringify(
      {
        date: TARGET_DATE,
        generatedAt: new Date().toISOString(),
        ordersCount: orders.length,
        channelSummary: state.channelSummary,
        totalRevenueCents,
        chipsSalesCents: state.chipsSalesCents,
        chipsSalesByChannel: state.chipsSalesByChannel,
        chipsPortionFromPacksQty: state.chipsPortionFromPacksQty,
      },
      null,
      2
    ),
    "utf8"
  );

  console.log(`\n已保存产品报表: ${itemsTxtPath} / ${itemsJsonPath}`);
  console.log(`已保存渠道报表: ${channelsTxtPath} / ${channelsJsonPath}`);
}

main().catch((err) => {
  console.error("脚本出错:", err);
  process.exit(1);
});
