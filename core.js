// ========== 扁平化教学大纲数据表（源自《机修大纲修订版2022.7.xlsx》，AI选行号，程序填课时） ==========
var DEFAULT_SYLLABUS_ROWS = [
    {"id": 0, "chapter": "汽车认识", "task": "汽车的诞生、发展及新国标分类", "skill": "", "content": "汽车的诞生、发展及新国标汽车分类方法", "total": 1.5, "theory": 1.5, "practical": 0.0},
    {"id": 1, "chapter": "汽车维修工安全知识", "task": "8S管理与维修工具设备安全规范", "skill": "", "content": "8S管理维修安全与工具设备安全使用", "total": 0.75, "theory": 0.75, "practical": 0.0},
    {"id": 2, "chapter": "常用工具及专用工具的了解与使用", "task": "常用工具与专用工具的规范使用", "skill": "", "content": "常用工具及专用工具的介绍与使用", "total": 2.25, "theory": 2.25, "practical": 0.0},
    {"id": 3, "chapter": "发动机的整体结构拆装认知", "task": "发动机的定义、分类及总体构造", "skill": "", "content": "发动机定义、分类及各大机构与系统组成作用", "total": 1.5, "theory": 1.5, "practical": 0.0},
    {"id": 4, "chapter": "发动机的整体结构拆装认知", "task": "发动机基本工作原理与专业术语", "skill": "", "content": "发动机四冲程工作原理及基本专业术语", "total": 3.0, "theory": 3.0, "practical": 0.0},
    {"id": 5, "chapter": "发动机的整体结构拆装认知", "task": "发动机总成外部附件拆装实训", "skill": "发动机拆装步骤", "content": "发动机各零部件安装位置及拆装步骤实训", "total": 12.0, "theory": 1.0, "practical": 11.0},
    {"id": 6, "chapter": "曲柄连杆机构的检测及装配", "task": "曲柄连杆机构的功用与组成认知", "skill": "曲柄连杆机构的认知", "content": "曲柄连杆机构的作用、组成及工作过程", "total": 1.5, "theory": 1.0, "practical": 0.5},
    {"id": 7, "chapter": "曲柄连杆机构的检测及装配", "task": "机体组的结构原理与分解装配", "skill": "机体组的装配", "content": "机体组各部件特点、分解组装步骤及检修", "total": 9.0, "theory": 1.0, "practical": 8.0},
    {"id": 8, "chapter": "曲柄连杆机构的检测及装配", "task": "活塞连杆组的结构与拆装检测", "skill": "活塞连杆组装配", "content": "活塞连杆组部件构造、活塞环安装与分解组装", "total": 9.0, "theory": 1.5, "practical": 7.5},
    {"id": 9, "chapter": "曲柄连杆机构的检测及装配", "task": "曲轴飞轮组的装配与常见故障诊断", "skill": "曲轴飞轮组的装配", "content": "曲轴飞轮组的分解组装及曲柄连杆机构故障排除", "total": 9.0, "theory": 1.5, "practical": 7.5},
    {"id": 10, "chapter": "配气机构的检测及维修", "task": "配气机构的作用、组成及配气相位", "skill": "配气机构整体认知", "content": "配气机构组成、凸轮轴布置传动方式及配气相位", "total": 6.0, "theory": 5.5, "practical": 0.5},
    {"id": 11, "chapter": "配气机构的检测及维修", "task": "气门组的结构与气门油封更换", "skill": "气门组件的更换", "content": "气门组的组成结构及气门、气门油封的更换", "total": 6.0, "theory": 1.0, "practical": 5.0},
    {"id": 12, "chapter": "配气机构的检测及维修", "task": "气门传动组的构造与拆装步骤", "skill": "", "content": "气门传动组的组成讲解及分解组装步骤", "total": 6.0, "theory": 6.0, "practical": 0.0},
    {"id": 13, "chapter": "配气机构的检测及维修", "task": "气门间隙的检查与调整实训", "skill": "气门间隙调整", "content": "气门间隙的定义、影响及调整方法实训", "total": 12.0, "theory": 1.0, "practical": 11.0},
    {"id": 14, "chapter": "配气机构的检测及维修", "task": "正时标记校对与正时皮带更换", "skill": "正时皮带更换", "content": "正时标记找取校对、正时皮带安装更换及检查", "total": 12.0, "theory": 2.0, "practical": 10.0},
    {"id": 15, "chapter": "发动机冷却系统", "task": "冷却系统的功用、组成与循环原理", "skill": "", "content": "冷却系的作用、组成及大小循环工作原理", "total": 1.5, "theory": 1.5, "practical": 0.0},
    {"id": 16, "chapter": "发动机冷却系统", "task": "散热器、水泵与节温器构造原理", "skill": "", "content": "散热器、水泵、节温器及风扇离合器结构原理", "total": 6.0, "theory": 6.0, "practical": 0.0},
    {"id": 17, "chapter": "发动机冷却系统", "task": "冷却系统水温过高的故障诊断排除", "skill": "", "content": "冷却系统水温过高的原因分析及故障排除", "total": 1.5, "theory": 1.5, "practical": 0.0},
    {"id": 18, "chapter": "发动机冷却系统", "task": "冷却系统泄漏检查与水管拆装更换", "skill": "冷却系统泄漏的检查及水管的更换", "content": "冷却系统泄漏常见部位检查及水管更换", "total": 4.5, "theory": 0.5, "practical": 4.0},
    {"id": 19, "chapter": "发动机冷却系统", "task": "防冻液的选用与更换操作实训", "skill": "防冻液的更换", "content": "防冻液的作用、选用及规范更换步骤", "total": 4.5, "theory": 0.5, "practical": 4.0},
    {"id": 20, "chapter": "发动机冷却系统", "task": "水泵与节温器的检查及更换实训", "skill": "水泵、节温器的检查与更换", "content": "水泵、节温器的检查与更换操作步骤", "total": 9.0, "theory": 1.0, "practical": 8.0},
    {"id": 21, "chapter": "发动机润滑系统", "task": "润滑系统的功用、油路与曲轴箱通风", "skill": "", "content": "润滑系统的作用、组成原理及曲轴箱通风", "total": 1.5, "theory": 1.5, "practical": 0.0},
    {"id": 22, "chapter": "发动机润滑系统", "task": "机油泵与机油滤清器的构造原理", "skill": "", "content": "机油泵及机油滤清器的构造与工作原理", "total": 3.0, "theory": 3.0, "practical": 0.0},
    {"id": 23, "chapter": "发动机润滑系统", "task": "机油及机油滤清器的更换整车实训", "skill": "机油及机油滤清器的更换（整车实训）", "content": "机油分类选用、液位检查及机油滤清器更换", "total": 12.0, "theory": 1.5, "practical": 10.5},
    {"id": 24, "chapter": "发动机润滑系统", "task": "机油泵的检查项目与更换步骤", "skill": "", "content": "机油泵的检查项目及更换步骤注意事项", "total": 6.0, "theory": 6.0, "practical": 0.0},
    {"id": 25, "chapter": "发动机总成维修", "task": "发动机总成分解、检修、装配与竣工验收", "skill": "分解、装配", "content": "发动机总成吊取、分解、清洗、检修与大修装配", "total": 30.0, "theory": 3.5, "practical": 26.5},
    {"id": 26, "chapter": "汽车底盘概述", "task": "汽车底盘的功用与四大系统组成", "skill": "汽车底盘整体认知", "content": "汽车底盘的发展史、功用与总体组成", "total": 1.5, "theory": 1.0, "practical": 0.5},
    {"id": 27, "chapter": "汽车底盘概述", "task": "汽车底盘的总体布置与驱动形式", "skill": "汽车底盘的驱动形式", "content": "汽车底盘常见传动布置与驱动形式", "total": 1.5, "theory": 1.0, "practical": 0.5},
    {"id": 28, "chapter": "离合器", "task": "离合器的功用、组成与工作原理", "skill": "离合器的认知", "content": "离合器的功用、组成、类型及工作过程", "total": 3.0, "theory": 0.5, "practical": 2.5},
    {"id": 29, "chapter": "离合器", "task": "离合器三件套的检查与整车更换实训", "skill": "离合器的检查及更换（整车实训）", "content": "从动盘、压盘、分离轴承的检查及更换流程", "total": 18.0, "theory": 1.0, "practical": 17.0},
    {"id": 30, "chapter": "离合器", "task": "离合器操纵机构的检查与自由行程调整", "skill": "操纵机构的检查及调整", "content": "离合器操纵机构类型、自由行程检查与调整", "total": 9.0, "theory": 2.0, "practical": 7.0},
    {"id": 31, "chapter": "离合器", "task": "离合器常见故障现象与诊断分析", "skill": "", "content": "离合器常见故障现象、原因及诊断流程", "total": 1.5, "theory": 1.5, "practical": 0.0},
    {"id": 32, "chapter": "手动变速器", "task": "变速器的功用、分类与安装位置", "skill": "", "content": "手动变速器的功能、安装位置及分类", "total": 2.0, "theory": 2.0, "practical": 0.0},
    {"id": 33, "chapter": "手动变速器", "task": "手动变速器变速机构与齿轮传动过程", "skill": "", "content": "变速器操作机构、变速机构及齿轮变速过程", "total": 4.0, "theory": 4.0, "practical": 0.0},
    {"id": 34, "chapter": "手动变速器", "task": "手动变速器各档动力传递与同步器原理", "skill": "", "content": "手动变速器工作原理与同步器啮合过程", "total": 3.0, "theory": 3.0, "practical": 0.0},
    {"id": 35, "chapter": "手动变速器", "task": "变速器齿轮油的选用与整车更换实训", "skill": "油液的更换（整车实训）", "content": "变速器齿轮油的作用、质量辨别及更换流程", "total": 6.0, "theory": 1.5, "practical": 4.5},
    {"id": 36, "chapter": "手动变速器", "task": "手动变速器总成的拆装与零部件检修", "skill": "变速器的拆装及检查", "content": "手动变速器拆装步骤及齿轮、同步器、轴承检修", "total": 18.0, "theory": 1.0, "practical": 17.0},
    {"id": 37, "chapter": "手动变速器", "task": "手动变速器跳挡、乱挡及异响故障诊断", "skill": "", "content": "变速器跳挡、乱挡、异响、漏油故障诊断", "total": 1.5, "theory": 1.5, "practical": 0.0},
    {"id": 38, "chapter": "手动变速器", "task": "变速器换挡操纵机构的检查与调整", "skill": "操纵机构的检查及调整", "content": "变速器操纵机构结构认知及档位检查调整", "total": 6.0, "theory": 1.0, "practical": 5.0},
    {"id": 39, "chapter": "驱动桥", "task": "驱动桥、主减速器与差速器的构造原理", "skill": "驱动桥的认知", "content": "驱动桥的功用、组成、类型及差速工作过程", "total": 1.5, "theory": 1.0, "practical": 0.5},
    {"id": 40, "chapter": "驱动桥", "task": "驱动桥常见故障现象与诊断分析", "skill": "", "content": "驱动桥常见故障点、异响原因及分析流程", "total": 1.5, "theory": 1.5, "practical": 0.0},
    {"id": 41, "chapter": "万向传动装置", "task": "万向节与传动轴的功用、类型及原理", "skill": "万向传动装置的认知", "content": "万向节与传动轴的作用、类型及工作过程", "total": 3.0, "theory": 2.5, "practical": 0.5},
    {"id": 42, "chapter": "万向传动装置", "task": "万向传动装置的检查与整车更换实训", "skill": "万向传动装置的检查与更换（整车实训）", "content": "万向节与传动轴的检查、更换及常见故障", "total": 9.0, "theory": 2.0, "practical": 7.0},
    {"id": 43, "chapter": "行驶系", "task": "行驶系的功用、组成与车架车桥结构", "skill": "行驶系的认知", "content": "行驶系的功用组成及车架、车桥类型与结构", "total": 6.0, "theory": 5.5, "practical": 0.5},
    {"id": 44, "chapter": "行驶系", "task": "车轮定位参数与四轮定位检测调整实训", "skill": "四轮定位的检查与调整（整车实训）", "content": "四轮定位的参数特点、专业术语及检查调整", "total": 9.0, "theory": 2.0, "practical": 7.0},
    {"id": 45, "chapter": "行驶系", "task": "车轮与轮胎构造、动平衡及扒胎补胎实训", "skill": "轮胎的检查与更换（整车实训）扒胎、补胎", "content": "轮胎规格表示、动平衡、轮胎换位及扒胎补胎", "total": 12.0, "theory": 1.5, "practical": 10.5},
    {"id": 46, "chapter": "行驶系", "task": "悬架系统的功用、组成与结构类型", "skill": "悬架系统的认知", "content": "悬架装置的功用、安装位置、组成及类型", "total": 3.0, "theory": 0.5, "practical": 2.5},
    {"id": 47, "chapter": "行驶系", "task": "减震器与弹性元件的检查及整车更换实训", "skill": "悬架的检查与更换（整车实训）", "content": "弹性元件与减震器的检查及减震器更换流程", "total": 12.0, "theory": 1.0, "practical": 11.0},
    {"id": 48, "chapter": "行驶系", "task": "行驶系常见故障现象与诊断排除", "skill": "", "content": "行驶系常见故障现象、原因及诊断流程", "total": 1.5, "theory": 1.5, "practical": 0.0},
    {"id": 49, "chapter": "转向系", "task": "转向系的功用、组成与转向工作原理", "skill": "转向系认知", "content": "转向系的功用、组成类别及转向工作过程", "total": 1.5, "theory": 1.0, "practical": 0.5},
    {"id": 50, "chapter": "转向系", "task": "转向球头与拉杆的检查及更换实训", "skill": "球头及拉杆的检查与更换（整车实训）", "content": "转向球头及拉杆的作用、检查与更换流程", "total": 6.0, "theory": 1.0, "practical": 5.0},
    {"id": 51, "chapter": "转向系", "task": "转向器的结构类型与检查更换流程", "skill": "", "content": "转向机的结构类型及检查更换流程", "total": 1.5, "theory": 1.5, "practical": 0.0},
    {"id": 52, "chapter": "转向系", "task": "转向助力泵与助力油的检查及更换实训", "skill": "转向助力泵及油液的更换（整车实训）", "content": "转向助力油、助力泵皮带及助力泵的检查与更换", "total": 12.0, "theory": 2.0, "practical": 10.0},
    {"id": 53, "chapter": "转向系", "task": "转向系常见故障现象与诊断分析", "skill": "", "content": "转向系常见故障现象、原因及诊断流程", "total": 1.5, "theory": 1.5, "practical": 0.0},
    {"id": 54, "chapter": "制动系统", "task": "制动系统的功用、组成与工作原理", "skill": "制动系的认知", "content": "制动系统的作用、组成、工作过程及类型", "total": 4.5, "theory": 4.0, "practical": 0.5},
    {"id": 55, "chapter": "制动系统", "task": "制动液、制动片与制动盘鼓的更换实训", "skill": "制动油液、制动片、制动鼓的检查及更换（整车实训）", "content": "制动液、制动片、制动盘鼓的检查及整车更换", "total": 18.0, "theory": 1.5, "practical": 16.5},
    {"id": 56, "chapter": "制动系统", "task": "真空助力器的组成结构与工作原理", "skill": "", "content": "真空助力泵的组成结构、工作过程与更换步骤", "total": 1.5, "theory": 1.5, "practical": 0.0},
    {"id": 57, "chapter": "制动系统", "task": "驻车制动系统的检查与调整实训", "skill": "驻车制动的检查与调整（整车实训）", "content": "驻车制动的组成结构及检查调整操作流程", "total": 6.0, "theory": 1.5, "practical": 4.5},
    {"id": 58, "chapter": "制动系统", "task": "轮毂轴承的检查、保养与更换流程", "skill": "轮毂轴承的保养、更换（整车实训）", "content": "轮毂轴承的作用、损坏原因及保养更换流程", "total": 1.5, "theory": 1.0, "practical": 0.5},
    {"id": 59, "chapter": "制动系统", "task": "制动系统常见故障现象与诊断分析", "skill": "", "content": "制动系常见故障现象、原因及诊断流程", "total": 1.5, "theory": 1.5, "practical": 0.0}
];

// ========== 全局常量与状态 ==========
var currentHost = null;
var draggedSource = null;
var CONFIG = {
    teacher: "",
    major: "",
    model: "",
    url: "",
    key: "",
    classes: [],
    courses: [],
    defaultView: "view-home",
    namingVals: [],
    namingSeps: [],
    excelNamingVals: [],
    excelNamingSeps: [],
    expectedStudents: {},
    excelDailySettingsMap: {},
    masterContentMap: {},
    customTemplateData: null,
    customSyllabusRows: []
};

var lastRootView = "view-home";
var lastFormData = null;
var lastAiData = null;
var uploadedImages = [];

function normalizeApiUrl(rawUrl) {
    if (!rawUrl || typeof rawUrl !== "string") return "";
    var trimmed = rawUrl.trim();
    if (!trimmed) return "";
    if (!/^https?:\/\//i.test(trimmed)) {
        trimmed = "https://" + trimmed;
    }
    return trimmed;
}

// 获取所有大纲数据行（内置 + 用户导入拓展）
function getAllSyllabusRows() {
    var custom = Array.isArray(CONFIG.customSyllabusRows) ? CONFIG.customSyllabusRows : [];
    var combined = DEFAULT_SYLLABUS_ROWS.concat(custom);
    return combined.map(function (row, idx) {
        return {
            id: idx,
            chapter: row.chapter || "",
            task: row.task || "",
            skill: row.skill || "",
            content: row.content || "",
            total: Number(row.total || 0),
            theory: Number(row.theory || 0),
            practical: Number(row.practical || 0)
        };
    });
}

// ========== 全局提示与状态管理 ==========
function showStatus(text, type) {
    type = type || "info";
    var el = document.getElementById("status-msg");
    var el2 = document.getElementById("excel-status-msg");
    var el3 = document.getElementById("plan-status-msg");
    if (el) el.innerHTML = "<span class='status-" + type + "'>" + text + "</span>";
    if (el2) el2.innerHTML = "<span class='status-" + type + "'>" + text + "</span>";
    if (el3) el3.innerHTML = "<span class='status-" + type + "'>" + text + "</span>";
}

function clearStatus() {
    var el1 = document.getElementById("status-msg");
    var el2 = document.getElementById("excel-status-msg");
    var el3 = document.getElementById("plan-status-msg");
    if (el1) el1.innerHTML = "";
    if (el2) el2.innerHTML = "";
    if (el3) el3.innerHTML = "";
}

function saveStorage() {
    if (CONFIG.url) CONFIG.url = normalizeApiUrl(CONFIG.url);
    localStorage.setItem("schema_config", JSON.stringify(CONFIG));
}

// 节假日智能抓取与本地缓存系统
async function getHolidayConfig(year) {
    var cacheKey = "schemaai_holidays_" + year;
    var cached = localStorage.getItem(cacheKey);

    if (cached) {
        try {
            return JSON.parse(cached);
        } catch (e) {
            localStorage.removeItem(cacheKey);
        }
    }

    try {
        var response = await fetch("https://api.apisbo.com/holidays/year/" + year);
        if (!response.ok) throw new Error("API 请求失败");
        var resData = await response.json();
        var holidayMap = {};

        if (resData && resData.code === 0 && Array.isArray(resData.data)) {
            resData.data.forEach(function (item) {
                var dParts = item.date.split('-');
                if (dParts.length === 3) {
                    var dStr = parseInt(dParts[0]) + "." + parseInt(dParts[1]) + "." + parseInt(dParts[2]);
                    holidayMap[dStr] = item.type;
                }
            });
        }

        localStorage.setItem(cacheKey, JSON.stringify(holidayMap));
        return holidayMap;
    } catch (error) {
        console.warn("节假日 API 同步失败，降级为常规双休模式", error);
        return {};
    }
}