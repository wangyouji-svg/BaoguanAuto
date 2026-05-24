import re

path = "/www/server/panel/vhost/nginx/relay_proxy.conf"
with open(path) as f:
    content = f.read()

block = (
    "    location /gen {\n"
    "        proxy_pass http://127.0.0.1:5000/gen;\n"
    "        proxy_set_header Host $http_host;\n"
    "        proxy_set_header X-Real-IP $remote_addr;\n"
    "        proxy_read_timeout 120s;\n"
    "    }"
)

content_new = re.sub(r"    location /gen \{[^}]+\}", block, content)
with open(path, "w") as f:
    f.write(content_new)
print("done")
