/**
 * 报关资料自动生成 - DingTalk 表格脚本（纯同步版）
 *
 * 使用方式：
 *   1. 在「报关资料」工作表的 A 列（从 A2 开始）逐行填写要生成报关单的「合同号码」
 *      （每行一个，如 PKCELL-2024-001；遇到空行停止读取）
 *   2. 在脚本面板点击「运行脚本」
 *   3. B 列自动写入「点击下载」超链接，点击即可在浏览器下载 Excel 报关单
 *
 * 「报关资料」工作表列格式（第1行为表头）：
 *   A: 合同号码（用户填写）  B: 下载链接  C: TraceId  D: 时间  E: 状态  F: 错误信息  G: 备注
 *
 * 主数据表列名要求（第一行为列头）：
 *   合同号码、境外收货人、贸易国、运抵国、成交方式
 *   商品编号、商品名称、品牌、规格型号、数量、单位、单价、金额、币别、货源地
 *
 * 运费规则：
 *   同一合同号码下，若某行“商品编号”为空且“商品名称”为“国际运费”，
 *   后端会把该行“单价+币别”写入报关单运费栏（K12，合并单元格主单元格）。
 *
 * 原理：脚本先把订单数据发到服务器侧 SQLite 临时缓存，再把短 token 写入下载链接。
 * 用户点击超链接时，服务器按 token 取回缓存数据并实时生成 Excel 下载。
 */

(function () {
    var DATA_SHEET_NAME = '报关数据';
    var RESULT_SHEET_NAME = '报关资料';
    var SCRIPT_VERSION = 'cache-v1';
    var CACHE_ENDPOINTS = [
        'https://pkcellsolution.com/baoguan/cache',
        'https://pkcellsolution.com/baoguan/generate?cache=1'
    ];

    var REQUIRED_COLUMNS = [
        '合同号码', '境外收货人', '贸易国', '运抵国', '成交方式',
        '商品编号', '商品名称', '品牌', '规格型号', '数量', '单位', '单价', '金额', '币别', '货源地',
        '箱数', '毛重', '净重'
    ];

    // 后端填充报关单所需的关键字段（前端即使抓不到也会补空字符串，避免流程中断）
    var BACKEND_FIELDS = [
        '合同号码', '境外收货人', '贸易国', '运抵国', '成交方式',
        '商品编号', '商品名称', '品牌', '规格型号', '数量', '单位', '单价', '金额', '币制', '货源地',
        '箱数', '毛重', '净重', '毛重（千克）', '净重（千克）'
    ];

    // 报关数据表中箱单联动字段固定列位（0-based）：W=22, X=23, Y=24
    var WXY_INDEX = {
        carton: 22,
        gross: 23,
        net: 24,
    };

    var FIELD_ALIASES = {
        '毛重': ['毛重', '毛重（千克）', '毛重(千克)', '毛重(KG)', '毛重kg'],
        '净重': ['净重', '净重（千克）', '净重(千克)', '净重(KG)', '净重kg']
    };

    function safeSheetName(sheet) {
        try {
            return sheet && sheet.getName ? String(sheet.getName()) : '';
        } catch (e) {
            return '';
        }
    }

    function makeTraceId(contractNo) {
        var now = new Date();
        var ts = now.getFullYear()
            + String(now.getMonth() + 1).padStart(2, '0')
            + String(now.getDate()).padStart(2, '0')
            + String(now.getHours()).padStart(2, '0')
            + String(now.getMinutes()).padStart(2, '0')
            + String(now.getSeconds()).padStart(2, '0');
        var rand = Math.random().toString(36).slice(2, 8).toUpperCase();
        var key = String(contractNo || '').replace(/[^A-Za-z0-9]/g, '').slice(-6).toUpperCase();
        return 'BG-' + ts + '-' + (key || 'NA') + '-' + rand;
    }

    function firstNonEmptyField(obj, aliases) {
        for (var i = 0; i < aliases.length; i++) {
            var key = aliases[i];
            var val = obj[key];
            if (val !== null && val !== undefined && String(val).trim() !== '') {
                return String(val);
            }
        }
        return '';
    }

    function postJsonSync(url, payload, traceId) {
        if (typeof XMLHttpRequest === 'undefined') {
            throw new Error('当前脚本环境不支持 XMLHttpRequest');
        }
        var reqUrl = url + (url.indexOf('?') >= 0 ? '&' : '?')
            + 'trace_id=' + encodeURIComponent(traceId);
        var xhr = new XMLHttpRequest();
        xhr.open('POST', reqUrl, false);
        xhr.setRequestHeader('Content-Type', 'application/json;charset=UTF-8');
        xhr.send(JSON.stringify(payload));
        var responseText = xhr.responseText || '';
        if (xhr.status < 200 || xhr.status >= 300) {
            throw new Error('HTTP ' + xhr.status + ': ' + responseText);
        }
        if (!responseText) {
            return {};
        }
        return JSON.parse(responseText);
    }

    function buildLegacyUrl(dataStr, traceId) {
        var encoded = b64EncodeUtf8(dataStr);
        return 'https://pkcellsolution.com/baoguan/generate?d=' + encoded
            + '&trace_id=' + encodeURIComponent(traceId)
            + '&sv=' + encodeURIComponent(SCRIPT_VERSION);
    }

    var activeSheet = Workbook.getActiveSheet();
    var dataSheet = Workbook.getSheet(DATA_SHEET_NAME);
    var resultSheet = Workbook.getSheet(RESULT_SHEET_NAME);

    if (!resultSheet) {
        if (activeSheet) {
            activeSheet.getRange('A1').setValue('错误：未找到「' + RESULT_SHEET_NAME + '」工作表');
        }
        return;
    }

    if (!dataSheet) {
        resultSheet.getRange('A1').setValue('错误：未找到数据源工作表「' + DATA_SHEET_NAME + '」');
        return;
    }

    // ---- 从「报关资料」sheet A 列读取合同号码（A2 起，0-based row index 1）----
    var rCount = resultSheet.getRowCount();
    var orderEntries = []; // [{rowIdx, contractNo}]
    for (var r = 1; r < rCount; r++) {
        var cell = resultSheet.getRange(r, 0, 1, 1).getValues()[0][0];
        var contractNo = String(cell || '').trim();
        if (!contractNo) break;
        orderEntries.push({ rowIdx: r, contractNo: contractNo });
    }

    if (orderEntries.length === 0) {
        resultSheet.getRange('A1').setValue('错误：请在「' + RESULT_SHEET_NAME + '」sheet 的 A2 起逐行填写合同号码');
        return;
    }

    var dataSheetName = safeSheetName(dataSheet);

    // ---- 读取主数据表全表（一次性）----
    var rowCount = dataSheet.getRowCount();
    var colCount = dataSheet.getColumnCount();
    var allValues = dataSheet.getRange(0, 0, rowCount, colCount).getValues();

    // ---- 解析列头（第0行）----
    var headers = allValues[0];
    var colMap = {};
    for (var hi = 0; hi < headers.length; hi++) {
        var h = headers[hi];
        if (h !== null && h !== undefined && String(h).trim() !== '') {
            colMap[String(h).trim()] = hi;
        }
    }
    var orderIdCol = colMap['合同号码'];

    var missingColumns = [];
    for (var ci = 0; ci < REQUIRED_COLUMNS.length; ci++) {
        var colName = REQUIRED_COLUMNS[ci];
        if (colMap[colName] === undefined) {
            missingColumns.push(colName);
        }
    }

    var globalRemark = '';
    if (missingColumns.length > 0) {
        globalRemark = '前端未获取到列：' + missingColumns.join('、')
            + '；请手动确认相关字段'
            + (dataSheetName ? '（数据表：' + dataSheetName + '）' : '');
    }

    // ---- 结果回填函数（写到「报关资料」同行 B-G 列）----
    // 列位：B 下载链接、C traceId、D 时间、E 状态、F 失败原因、G 备注
    function writeResult(rowIdx, url, status, errMsg, remarkMsg, traceId) {
        var now = new Date();
        var timeStr = now.getFullYear() + '-'
            + String(now.getMonth() + 1).padStart(2, '0') + '-'
            + String(now.getDate()).padStart(2, '0') + ' '
            + String(now.getHours()).padStart(2, '0') + ':'
            + String(now.getMinutes()).padStart(2, '0');

        var finalErr = errMsg || '';
        var finalRemark = remarkMsg || '';
        var finalTraceId = traceId || '';
        resultSheet.getRange(rowIdx, 1, 1, 6).setValues([[url ? '点击下载' : '', finalTraceId, timeStr, status, finalErr, finalRemark]]);

        if (url) {
            try {
                resultSheet.getCell(rowIdx, 1).setHyperlink({
                    type: 'path',
                    link: url,
                    text: '点击下载',
                });
            } catch (e) {
                resultSheet.getRange(rowIdx, 4, 1, 2).setValues([[
                    '失败',
                    '写入超链接失败（长度=' + url.length + '）: ' + (e && e.message ? e.message : '未知错误'),
                ]]);
            }
        }
    }

    function collectManualCheckFields(rows) {
        // 全局检查不再强制净重，净重仅在指定电池类型下做逐行提示。
        var fields = ['境外收货人', '合同号码', '贸易国', '运抵国', '成交方式', '商品编号', '商品名称', '规格型号', '数量', '单位', '单价', '金额', '币制', '货源地'];
        var missing = [];
        for (var fi = 0; fi < fields.length; fi++) {
            var field = fields[fi];
            var hasValue = false;
            for (var ri = 0; ri < rows.length; ri++) {
                var v = rows[ri][field];
                if (v !== null && v !== undefined && String(v).trim() !== '') {
                    hasValue = true;
                    break;
                }
            }
            if (!hasValue) {
                missing.push(field);
            }
        }
        return missing;
    }

    function isEmptyValue(v) {
        return v === null || v === undefined || String(v).trim() === '';
    }

    function collectPerRowMissingNotes(rows) {
        var notes = [];
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i] || {};
            var hsCode = String(row['商品编号'] || '').trim();
            var itemName = String(row['商品名称'] || '').trim();

            // 运费行不参与商品缺失项提示。
            if (!hsCode && itemName === '国际运费') {
                continue;
            }

            var missing = [];
            if (isEmptyValue(row['数量'])) {
                missing.push('数量');
            }
            if (hsCode === '8507600090' || hsCode === '8506500011') {
                var netValue = firstNonEmptyField(row, FIELD_ALIASES['净重']);
                if (isEmptyValue(netValue)) {
                    missing.push('净重');
                }
            }

            if (missing.length > 0) {
                notes.push('第' + (i + 1) + '条缺少：' + missing.join('、'));
            }
        }
        return notes;
    }

    function collectManualCheckNotes(rows) {
        return collectPerRowMissingNotes(rows);
    }

    // UTF-8 安全的 URL-safe base64 编码（支持中文字段值）
    function b64EncodeUtf8(str) {
        return btoa(unescape(encodeURIComponent(str)))
            .replace(/\+/g, '-')
            .replace(/\//g, '_')
            .replace(/=/g, '');
    }

    // ---- 按合同号码逐个处理 ----
    for (var oi = 0; oi < orderEntries.length; oi++) {
        var entry = orderEntries[oi];
        var targetOrderId = entry.contractNo;
        var targetRowIdx = entry.rowIdx;
        var traceId = makeTraceId(targetOrderId);

        // 筛选匹配当前合同号码的数据行
        var matchedRows = [];
        for (var r = 1; r < allValues.length; r++) {
            var row = allValues[r];
            var rowOrderId = orderIdCol !== undefined
                ? String(row[orderIdCol] || '').trim()
                : '';
            if (rowOrderId !== targetOrderId) continue;

            var obj = {};
            for (var key in colMap) {
                var val = row[colMap[key]];
                obj[key] = (val !== null && val !== undefined) ? String(val) : '';
            }

            // 列名别名归一：优先取任一非空值，统一写入标准字段。
            var normalizedGross = firstNonEmptyField(obj, FIELD_ALIASES['毛重']);
            var normalizedNet = firstNonEmptyField(obj, FIELD_ALIASES['净重']);
            if (normalizedGross) {
                obj['毛重'] = normalizedGross;
            }
            if (normalizedNet) {
                obj['净重'] = normalizedNet;
            }

            // 兜底：无论列头命名如何，直接按 W/X/Y 列位抓取箱单联动字段。
            if (!obj['箱数']) {
                var wVal = row[WXY_INDEX.carton];
                obj['箱数'] = (wVal !== null && wVal !== undefined) ? String(wVal) : '';
            }
            if (!obj['毛重']) {
                var xVal = row[WXY_INDEX.gross];
                obj['毛重'] = (xVal !== null && xVal !== undefined) ? String(xVal) : '';
            }
            if (!obj['净重']) {
                var yVal = row[WXY_INDEX.net];
                obj['净重'] = (yVal !== null && yVal !== undefined) ? String(yVal) : '';
            }

            // 兼容后端历史字段名。
            if (!obj['毛重（千克）'] && obj['毛重']) {
                obj['毛重（千克）'] = obj['毛重'];
            }
            if (!obj['净重（千克）'] && obj['净重']) {
                obj['净重（千克）'] = obj['净重'];
            }

            // 兼容后端字段名：源表使用“币别”，后端标准字段使用“币制”。
            if (!obj['币制'] && obj['币别']) {
                obj['币制'] = obj['币别'];
            }

            // 非阻断模式：即使前端抓不到某些列，也补齐关键字段为空字符串，继续执行。
            for (var bi = 0; bi < BACKEND_FIELDS.length; bi++) {
                var bk = BACKEND_FIELDS[bi];
                if (obj[bk] === undefined) {
                    obj[bk] = '';
                }
            }

            matchedRows.push(obj);
        }

        if (matchedRows.length === 0) {
            writeResult(
                targetRowIdx,
                '',
                '待确认',
                '',
                '前端未查询到合同号码为「' + targetOrderId + '」的数据；请手动确认该合同号码是否存在于「' + DATA_SHEET_NAME + '」',
                traceId
            );
            continue;
        }

        var remarks = [];
        if (globalRemark) {
            remarks.push(globalRemark);
        }
        var needManualFields = collectManualCheckFields(matchedRows);
        if (needManualFields.length > 0) {
            remarks.push('以下字段无有效值：' + needManualFields.join('、') + '；请手动确认');
        }
        var manualNotes = collectManualCheckNotes(matchedRows);
        if (manualNotes.length > 0) {
            remarks.push(manualNotes.join('；'));
        }
        var finalRemark = remarks.join('；');

        var dataStr = JSON.stringify({ rows: matchedRows, meta: { traceId: traceId, scriptVersion: SCRIPT_VERSION } });
        try {
            var cacheResp = null;
            var lastCacheError = '';
            for (var ei = 0; ei < CACHE_ENDPOINTS.length; ei++) {
                try {
                    cacheResp = postJsonSync(CACHE_ENDPOINTS[ei], {
                        rows: matchedRows,
                        meta: { traceId: traceId, scriptVersion: SCRIPT_VERSION, cacheOnly: true }
                    }, traceId);
                    break;
                } catch (innerErr) {
                    lastCacheError = String(innerErr && innerErr.message ? innerErr.message : innerErr);
                }
            }
            if (!cacheResp) {
                throw new Error(lastCacheError || '缓存接口请求失败');
            }
            var shortUrl = cacheResp && cacheResp.url ? String(cacheResp.url) : '';
            if (!shortUrl) {
                var token = cacheResp && cacheResp.token ? String(cacheResp.token) : '';
                if (!token) {
                    throw new Error('缓存接口未返回 token');
                }
                shortUrl = 'https://pkcellsolution.com/baoguan/generate?t=' + encodeURIComponent(token)
                    + '&trace_id=' + encodeURIComponent(traceId);
            }
            writeResult(targetRowIdx, shortUrl, '链接已生成', '', finalRemark, traceId);
        } catch (cacheErr) {
            var legacyUrl = buildLegacyUrl(dataStr, traceId);
            if (legacyUrl.length <= 18000) {
                writeResult(targetRowIdx, legacyUrl, '链接已生成', '', finalRemark, traceId);
            } else {
                writeResult(targetRowIdx, '', '缓存失败', String(cacheErr && cacheErr.message ? cacheErr.message : cacheErr), finalRemark, traceId);
            }
        }
    }
})();
