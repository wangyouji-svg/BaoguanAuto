"""临时修复脚本：把 backend_server.py 里 '../' 相对路径改成同级目录"""
target = '/root/baoguan-backend/backend_server.py'
with open(target, 'r', encoding='utf-8') as f:
    c = f.read()

c = c.replace(
    "os.path.join(os.path.dirname(__file__), '..', '报关资料模板.xlsx')",
    "os.path.join(os.path.dirname(__file__), '报关资料模板.xlsx')"
)
c = c.replace(
    "os.path.join(os.path.dirname(__file__), '..', 'generated')",
    "os.path.join(os.path.dirname(__file__), 'generated')"
)

with open(target, 'w', encoding='utf-8') as f:
    f.write(c)
print('path fix ok')
