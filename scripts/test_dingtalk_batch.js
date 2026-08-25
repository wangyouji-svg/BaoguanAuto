const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const contractCount = 130;
const headers = [
    '合同号码', '境外收货人', '贸易国', '运抵国', '成交方式',
    '商品编号', '商品名称', '品牌', '规格型号', '数量', '单位',
    '单价', '金额', '币别', '货源地', '箱数', '毛重', '净重', '系统号码',
    '', '', '', '', '', ''
];
const dataRows = [headers];
const contracts = [];
for (let index = 0; index < contractCount; index += 1) {
    const contractNo = `PK-BATCH-${String(index + 1).padStart(3, '0')}`;
    contracts.push(contractNo);
    dataRows.push([
        contractNo, 'Receiver', '美国', '美国', 'FOB',
        '8507600090', '锂离子电池', 'PKCELL', 'IFR14500-800-3.2V',
        '10', 'PCS', '1.23', '12.30', 'USD', '深圳',
        '1', '2.0', '1.8', `SHIP-${index + 1}`
    ]);
}

const writtenRows = new Set();
const hyperlinks = new Set();
const resultSheet = {
    getRowCount() {
        return contractCount + 1;
    },
    getRange(row, column) {
        if (typeof row === 'string') {
            return { setValue() {} };
        }
        return {
            getValues() {
                return [[column === 0 && row > 0 ? contracts[row - 1] : '']];
            },
            setValues() {
                writtenRows.add(row);
            },
            setValue() {},
        };
    },
    getCell(row) {
        return {
            setHyperlink(link) {
                assert.match(link.link, /^https:\/\/pkcellsolution\.com\/baoguan\/generate\?t=/);
                hyperlinks.add(row);
            },
        };
    },
};

const dataSheet = {
    getName() {
        return '报关数据';
    },
    getRowCount() {
        return dataRows.length;
    },
    getColumnCount() {
        return headers.length;
    },
    getRange() {
        return { getValues: () => dataRows };
    },
};

global.Workbook = {
    getActiveSheet: () => resultSheet,
    getSheet(name) {
        if (name === '报关数据') return dataSheet;
        if (name === '报关资料') return resultSheet;
        return null;
    },
};

const requests = [];
global.XMLHttpRequest = function XMLHttpRequest() {
    this.status = 0;
    this.responseText = '';
    this.open = (method, url, async) => {
        this.method = method;
        this.url = url;
        this.async = async;
    };
    this.setRequestHeader = () => {};
    this.send = (body) => {
        const payload = JSON.parse(body);
        requests.push({ method: this.method, url: this.url, async: this.async, payload });
        assert.match(this.url, /\/generate\?cache=1&batch=1/);
        this.status = 200;
        this.responseText = JSON.stringify({
            items: payload.items.map((item, index) => ({
                ok: true,
                token: `token-${index}`,
                url: `https://pkcellsolution.com/baoguan/generate?t=token-${index}`,
                traceId: item.meta.traceId,
            })),
            successCount: payload.items.length,
            failureCount: 0,
        });
    };
};

const scriptPath = path.join(__dirname, 'dingtalk_demo.js');
vm.runInThisContext(fs.readFileSync(scriptPath, 'utf8'), { filename: scriptPath });

assert.strictEqual(requests.length, 1, '130 个合同应合并为一次 HTTP 请求');
assert.strictEqual(requests[0].method, 'POST');
assert.strictEqual(requests[0].async, false, '钉钉脚本仍使用同步 XHR，但只同步等待一次');
assert.strictEqual(requests[0].payload.items.length, contractCount);
assert.strictEqual(writtenRows.size, contractCount);
assert.strictEqual(hyperlinks.size, contractCount);
console.log(`dingtalk batch ok: ${contractCount} contracts, ${requests.length} request`);
