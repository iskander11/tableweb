#!/bin/bash
# Run on VPS to pull latest code and restart

set -e
cd /var/www/tableweb

git pull origin main

cd app/frontend
npm install
npm run build

cd ../backend
npm install

pm2 restart tableweb-backend
echo "Update complete!"
