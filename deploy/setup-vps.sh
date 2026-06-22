#!/bin/bash
# Run this script on the VPS as root to set up the server

set -e

echo "=== Installing Node.js 20 ==="
curl -fsSL https://deb.nodesource.com/setup_20.x | bash -
apt-get install -y nodejs

echo "=== Installing PostgreSQL ==="
apt-get install -y postgresql postgresql-contrib

echo "=== Installing Nginx ==="
apt-get install -y nginx

echo "=== Installing PM2 ==="
npm install -g pm2

echo "=== Creating PostgreSQL database ==="
sudo -u postgres psql -c "CREATE USER tableweb WITH PASSWORD 'CHANGE_THIS_PASSWORD';"
sudo -u postgres psql -c "CREATE DATABASE tableweb OWNER tableweb;"
sudo -u postgres psql -c "GRANT ALL PRIVILEGES ON DATABASE tableweb TO tableweb;"

echo "=== Cloning repository ==="
cd /var/www
git clone https://iskander11:TOKEN_HERE@github.com/iskander11/tableweb.git
cd tableweb/app/backend

echo "=== Creating .env ==="
cat > .env << 'EOF'
PORT=3001
JWT_SECRET=CHANGE_THIS_TO_RANDOM_SECRET
DB_HOST=localhost
DB_PORT=5432
DB_NAME=tableweb
DB_USER=tableweb
DB_PASSWORD=CHANGE_THIS_PASSWORD
FRONTEND_URL=http://YOUR_DOMAIN_OR_IP
BACKUP_DIR=/var/www/tableweb/backups
EOF

echo "=== Installing backend dependencies ==="
npm install

echo "=== Running DB schema ==="
sudo -u postgres psql -d tableweb -f src/db/schema.sql

echo "=== Creating first admin user ==="
node -e "
import bcrypt from 'bcryptjs';
import pg from 'pg';
import dotenv from 'dotenv';
dotenv.config();
const pool = new pg.Pool({ host:'localhost',database:'tableweb',user:'tableweb',password:process.env.DB_PASSWORD });
const hash = await bcrypt.hash('admin123', 12);
await pool.query(\"INSERT INTO users (username,email,password_hash,role) VALUES ('admin','admin@tableweb.ru',\$1,'admin')\", [hash]);
console.log('Admin created: login=admin password=admin123');
pool.end();
" --input-type=module

echo "=== Building frontend ==="
cd /var/www/tableweb/app/frontend
npm install
npm run build

echo "=== Starting backend with PM2 ==="
cd /var/www/tableweb/app/backend
pm2 start src/index.js --name tableweb-backend
pm2 save
pm2 startup

echo "=== Configuring Nginx ==="
cat > /etc/nginx/sites-available/tableweb << 'NGINX'
server {
    listen 80;
    server_name _;

    # Frontend
    root /var/www/tableweb/app/frontend/dist;
    index index.html;

    location / {
        try_files $uri $uri/ /index.html;
    }

    # Backend API
    location /api/ {
        proxy_pass http://localhost:3001;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_cache_bypass $http_upgrade;
        # Stream Server-Sent Events (import progress) without buffering so the
        # progress bar animates live instead of arriving all at once.
        proxy_buffering off;
        proxy_cache off;
        proxy_read_timeout 3600s;
    }

    # WebSocket
    location /socket.io/ {
        proxy_pass http://localhost:3001;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "Upgrade";
        proxy_set_header Host $host;
    }
}
NGINX

ln -sf /etc/nginx/sites-available/tableweb /etc/nginx/sites-enabled/
rm -f /etc/nginx/sites-enabled/default
nginx -t && systemctl reload nginx

echo ""
echo "=== SETUP COMPLETE ==="
echo "Open http://<YOUR_SERVER_IP> in browser"
echo "Login: admin / admin123 (CHANGE AFTER FIRST LOGIN)"
