#!/bin/bash
# Start Flask server + public tunnel
cd "$(dirname "$0")"
echo "Starting server..."
python3 server.py &
SERVER_PID=$!
sleep 2
echo "Starting tunnel..."
ssh -o StrictHostKeyChecking=no -o ServerAliveInterval=60 -R 80:localhost:8888 nokey@localhost.run 2>&1 | while read line; do
  echo "$line"
  if echo "$line" | grep -q "tunneled with tls termination"; then
    URL=$(echo "$line" | grep -o 'https://[a-z0-9.]*\.lhr\.life')
    echo "PUBLIC_URL=$URL" > tunnel_url.txt
    echo "Public URL: $URL"
  fi
done
