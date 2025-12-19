#!/bin/bash

# 檢查是否以 root 權限執行
if [ "$EUID" -ne 0 ]; then 
  echo "請使用 sudo 或 root 權限執行此腳本。"
  exit 1
fi

# 設定變數
TIMESTAMP=$(date +%Y%m%d_%H%M%S)
HOSTNAME=$(hostname)
EXPORT_DIR="forensics_${HOSTNAME}_${TIMESTAMP}"
LOG_FILE="${EXPORT_DIR}/collection_log.txt"

# 建立輸出目錄
mkdir -p "$EXPORT_DIR"
echo "--- Linux 鑑識收集開始: $(date) ---" | tee -a "$LOG_FILE"

# 1. 系統基本資訊
echo "[*] 正在收集系統資訊..." | tee -a "$LOG_FILE"
uname -a > "${EXPORT_DIR}/system_info.txt"
uptime >> "${EXPORT_DIR}/system_info.txt"
ip addr >> "${EXPORT_DIR}/network_info.txt"

# 2. 收集 /var/log (包含 auth.log, syslog, messages 等)
echo "[*] 正在收集 /var/log..." | tee -a "$LOG_FILE"
mkdir -p "${EXPORT_DIR}/logs"
cp -r /var/log/* "${EXPORT_DIR}/logs/" 2>/dev/null

# 3. 收集 Cron 排程工作
echo "[*] 正在收集 Cron 排程..." | tee -a "$LOG_FILE"
mkdir -p "${EXPORT_DIR}/cron"
cp -r /var/spool/cron/crontabs "${EXPORT_DIR}/cron/user_crons" 2>/dev/null
cp /etc/crontab "${EXPORT_DIR}/cron/etc_crontab" 2>/dev/null
cp -r /etc/cron.* "${EXPORT_DIR}/cron/" 2>/dev/null

# 4. 收集所有使用者的 Bash History
echo "[*] 正在收集 Bash 歷史紀錄..." | tee -a "$LOG_FILE"
mkdir -p "${EXPORT_DIR}/bash_history"
# 收集 root 的
cp /root/.bash_history "${EXPORT_DIR}/bash_history/root_history" 2>/dev/null
# 遍歷 /home 下的所有使用者
for user_dir in /home/*; do
    username=$(basename "$user_dir")
    if [ -f "$user_dir/.bash_history" ]; then
        cp "$user_dir/.bash_history" "${EXPORT_DIR}/bash_history/${username}_history" 2>/dev/null
    fi
done

# 5. 收集 SSH 與 Sudoers 設定
echo "[*] 正在收集 SSH 與 Sudoers 配置..." | tee -a "$LOG_FILE"
mkdir -p "${EXPORT_DIR}/config"
cp -r /etc/ssh/sshd_config "${EXPORT_DIR}/config/sshd_config" 2>/dev/null
cp -r /etc/sudoers "${EXPORT_DIR}/config/sudoers" 2>/dev/null
cp -r /etc/sudoers.d "${EXPORT_DIR}/config/sudoers.d" 2>/dev/null

# 6. 額外分析：列出當前登入使用者與登入紀錄 (last, lastlog)
echo "[*] 正在分析登入紀錄..." | tee -a "$LOG_FILE"
last > "${EXPORT_DIR}/last_login.txt"
lastlog > "${EXPORT_DIR}/lastlog.txt"
who > "${EXPORT_DIR}/current_users.txt"

# 打包壓縮
echo "[*] 正在打包資料..." | tee -a "$LOG_FILE"
tar -czf "${EXPORT_DIR}.tar.gz" "$EXPORT_DIR"

echo "--- 收集完成！ ---" | tee -a "$LOG_FILE"
echo "結果存放在: ${EXPORT_DIR}.tar.gz"
