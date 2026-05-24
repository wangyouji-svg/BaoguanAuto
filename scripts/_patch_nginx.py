"""在 relay_proxy.conf 的 'location / {' 之前插入报关后端路由"""
conf = '/www/server/panel/vhost/nginx/relay_proxy.conf'
insert = """
    location /generate {
        proxy_pass http://127.0.0.1:5000/generate;
        proxy_set_header Host $http_host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_read_timeout 120s;
        add_header Access-Control-Allow-Origin *;
        add_header Access-Control-Allow-Headers "Content-Type,bypass-tunnel-reminder";
    }

    location /download/ {
        proxy_pass http://127.0.0.1:5000/download/;
        proxy_set_header Host $http_host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_read_timeout 120s;
    }

"""
with open(conf, 'r') as f:
    content = f.read()

if 'location /generate' in content:
    print('already patched, skip')
else:
    content = content.replace('    location / {', insert + '    location / {', 1)
    with open(conf, 'w') as f:
        f.write(content)
    print('nginx config patched ok')
