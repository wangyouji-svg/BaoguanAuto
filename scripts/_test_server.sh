#!/bin/bash
cd /root/baoguan-backend
echo "=== Python syntax check ==="
/root/logiflow-tracker/.venv/bin/python -m py_compile backend_server.py && echo "Syntax OK" || echo "Syntax ERROR"

echo ""
echo "=== Test empty request ==="
/root/logiflow-tracker/.venv/bin/python -c "
import sys
sys.path.insert(0, '.')
from backend_server import fill_template
result = fill_template([])
print('fill_template([]) OK, file:', result)
" 2>&1

echo ""
echo "=== Test one row ==="
/root/logiflow-tracker/.venv/bin/python -c "
import sys
sys.path.insert(0, '.')
from backend_server import fill_template, build_product_name
row = {
    '商品编号': '8506101210',
    '品牌': 'PKCELL',
    '规格型号': 'PKCELL-LR6-1800-1.5V-T',
    '数量': '4032',
    '单位': '个',
    '单价': '0.1075',
    '币制': '美元',
    '货源地': '东莞其他',
    '贸易国': '阿根廷',
    '合同号码': 'PK-TEST',
    '境外收货人': 'ACME CORP',
    '成交方式': 'FOB',
    '净重': '22176',
}
spec = build_product_name(row)
print('spec string:', spec)
result = fill_template([row])
print('fill_template OK, file:', result)
" 2>&1
