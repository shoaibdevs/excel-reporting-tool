sudo nano /etc/systemd/system/reports.socket
sudo nano /etc/systemd/system/reports.service

[Unit]
Description=reports daemon
Requires=reports.socket
After=network.target

[Service]
User=sammy
Group=www-data
WorkingDirectory=/home/sammy/myprojectdir
ExecStart=/home/softlance/reports/env/bin/gunicorn \
          --access-logfile - \
          --workers 3 \
          --bind unix:/run/reports.sock \
          reports.wsgi:application

[Install]
WantedBy=multi-user.target

sudo systemctl start reports.socket

sudo systemctl enable reports.socket

curl --unix-socket /run/reports.sock localhost
sudo systemctl status reports.socket

sudo systemctl status reports.socket

sudo systemctl daemon-reload
sudo systemctl restart reports.socket reports.service
sudo nginx -t && sudo systemctl restart nginx
