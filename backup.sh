#!/bin/bash
DB=/opt/midea-dashboard/dashboard.db
BACKUP_DIR=/opt/midea-dashboard/backups
RETENTION_DAYS=30

mkdir -p $BACKUP_DIR
DATE=$(date +%Y%m%d)
BACKUP_FILE=$BACKUP_DIR/dashboard_$DATE.db

cp $DB $BACKUP_FILE
gzip -f $BACKUP_FILE

SIZE=$(ls -lh $BACKUP_FILE.gz | awk '{print $5}')
SUMMARY=$(python3 -c "
import sqlite3
db=sqlite3.connect('$DB')
d=db.execute('SELECT COUNT(*) FROM daily_records').fetchone()[0]
a=db.execute('SELECT COUNT(*) FROM abnormal_records WHERE deleted=0').fetchone()[0]
print(f'daily={d} abnormal={a}')
")
echo "[$(date '+%Y-%m-%d %H:%M:%S')] $BACKUP_FILE.gz ($SIZE) | $SUMMARY" >> $BACKUP_DIR/backup.log

find $BACKUP_DIR -name 'dashboard_*.db.gz' -mtime +$RETENTION_DAYS -delete
