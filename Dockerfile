# FROM nginx:alpine

# # ใส่ไฟล์ default.conf เข้าไปแทน default config
# COPY default.conf /etc/nginx/conf.d/default.conf

# # คัดลอก HTML เข้าไปใน nginx root
# COPY HTML /usr/share/nginx/html/HTML
# COPY HTML /usr/share/nginx/html/HTML
# COPY CSS /usr/share/nginx/html/CSS
# COPY JS /usr/share/nginx/html/JS

FROM nginx:alpine

COPY default.conf /etc/nginx/conf.d/default.conf

COPY HTML/ /usr/share/nginx/html/
COPY CSS/ /usr/share/nginx/html/CSS/
COPY JS/ /usr/share/nginx/html/JS/
