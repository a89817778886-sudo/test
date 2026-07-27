#!/bin/bash
# =============================================================================
# install_auto_watcher.sh — АВТОМАТИЧЕСКАЯ заливка новых zip → GitHub
#
# Watcher в фоне следит за ~/Downloads. Как только там появляется новый
# lks-kp-app*.zip — сам распаковывает, коммитит и пушит в GitHub.
# Без brew, без fswatch — только встроенные утилиты macOS.
#
# Как пользоваться:
#   1. Запусти этот скрипт один раз:
#        bash ~/Downloads/lks-kp-app/install_auto_watcher.sh
#   2. Скрипт попросит GitHub Token (если ещё не введён)
#   3. Готово — теперь каждый мой новый zip автоматически улетает в GitHub
#      сразу после скачивания. Ничего не нужно делать.
#
# Как остановить: bash ~/Downloads/lks-kp-app/install_auto_watcher.sh --stop
# Как проверить лог: cat ~/rolls_watcher.log
# =============================================================================

set -e

REPO_URL="github.com/a89817778886-sudo/-.git"
REPO_DIR="$HOME/rolls-kran"
GITHUB_USER="a89817778886-sudo"
WATCHER_SCRIPT="$HOME/rolls_watcher.sh"
PLIST_FILE="$HOME/Library/LaunchAgents/com.rolls.watcher.plist"
LOG_FILE="$HOME/rolls_watcher.log"

# --- Остановка (если передан флаг --stop) ---
if [ "$1" = "--stop" ]; then
    echo "🛑 Останавливаю watcher..."
    launchctl unload "$PLIST_FILE" 2>/dev/null || true
    echo "✅ Watcher остановлен. Чтобы запустить снова — запусти этот скрипт без флагов."
    exit 0
fi

echo "🚀 Настройка автоматической заливки zip → GitHub"
echo ""

# --- 1. Проверяем что репо есть, если нет — клонируем ---
if [ -d "$REPO_DIR/.git" ]; then
    echo "✅ Репозиторий уже склонирован: $REPO_DIR"
else
    echo "📥 Репозиторий не найден. Нужен GitHub Personal Access Token."
    echo ""
    echo "   Если у тебя ещё нет токена:"
    echo "   1. Открой https://github.com/settings/tokens"
    echo "   2. Generate new token (classic) → отметь 'repo' → Generate"
    echo "   3. Скопируй токен (начинается с ghp_...)"
    echo ""
    read -sp "🔑 Вставь сюда токен (символы не показываются — это нормально): " GH_TOKEN
    echo ""
    if [ -z "$GH_TOKEN" ]; then
        echo "❌ Токен не введён. Отмена."
        exit 1
    fi
    echo "📥 Клонирую репозиторий..."
    git clone "https://${GITHUB_USER}:${GH_TOKEN}@${REPO_URL}" "$REPO_DIR"
    cd "$REPO_DIR"
    git config credential.helper osxkeychain 2>/dev/null || true
    echo "✅ Репозиторий склонирован в $REPO_DIR"
fi

# --- 2. Создаём watcher-скрипт ---
cat > "$WATCHER_SCRIPT" << 'EOF'
#!/bin/bash
# Watcher: следит за ~/Downloads и загружает новые zip в GitHub.
# Использует stat -f%m (mtime) — без внешних зависимостей.

DOWNLOADS="$HOME/Downloads"
REPO_DIR="$HOME/rolls-kran"
LOG="$HOME/rolls_watcher.log"
STATE="$HOME/.rolls_watcher_state"

echo "🚀 Watcher запущен $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG"
echo "   Слежу за: $DOWNLOADS/lks-kp-app*.zip" >> "$LOG"

# Читаем последний обработанный mtime (если есть)
if [ -f "$STATE" ]; then
    LAST_MTIME=$(cat "$STATE")
else
    # При первом запуске считаем что все текущие файлы уже обработаны
    LAST_MTIME=$(date +%s)
    echo "$LAST_MTIME" > "$STATE"
fi

while true; do
    # Ищем последний zip в Downloads с mtime > LAST_MTIME
    LATEST_ZIP=""
    LATEST_MTIME=0
    for f in "$DOWNLOADS"/lks-kp-app*.zip; do
        [ -f "$f" ] || continue
        MTIME=$(stat -f%m "$f" 2>/dev/null || echo 0)
        if [ "$MTIME" -gt "$LAST_MTIME" ] && [ "$MTIME" -gt "$LATEST_MTIME" ]; then
            LATEST_ZIP="$f"
            LATEST_MTIME=$MTIME
        fi
    done

    if [ -n "$LATEST_ZIP" ]; then
        # Ждём 3 секунды — вдруг файл ещё дописывается
        sleep 3
        # Проверяем что размер не 0
        SIZE=$(stat -f%z "$LATEST_ZIP" 2>/dev/null || echo 0)
        if [ "$SIZE" -lt 100000 ]; then
            echo "$(date '+%H:%M:%S') ⚠️ Файл ещё не докачался ($SIZE байт), жду..." >> "$LOG"
            sleep 5
            continue
        fi

        echo "" >> "$LOG"
        echo "$(date '+%Y-%m-%d %H:%M:%S') 📦 Новый zip: $LATEST_ZIP" >> "$LOG"

        cd "$REPO_DIR" || { echo "❌ Нет папки $REPO_DIR" >> "$LOG"; sleep 3; continue; }

        # Распаковываем
        rm -rf /tmp/lks-new
        mkdir -p /tmp/lks-new
        unzip -qo "$LATEST_ZIP" -d /tmp/lks-new 2>> "$LOG"

        # Синхронизируем
        rsync -a --delete \
            --exclude='.git' \
            --exclude='crm.db' \
            --exclude='crm_files' \
            --exclude='crm_backups' \
            --exclude='.dadata_token' \
            --exclude='history_memory.json' \
            --exclude='install_tkp.sh' \
            --exclude='install_auto_watcher.sh' \
            /tmp/lks-new/lks-kp-app/ ./ 2>> "$LOG"

        git add -A 2>> "$LOG"
        if git diff --cached --quiet; then
            echo "$(date '+%H:%M:%S') ⚠️ Нет изменений — код уже актуален" >> "$LOG"
            osascript -e 'display notification "Нет изменений — код уже актуален" with title "⚠️ ROLLS KRAN" sound name "Pop"' 2>/dev/null || true
        else
            MSG="auto $(date '+%d.%m.%Y %H:%M')"
            if git commit -q -m "$MSG" 2>> "$LOG"; then
                if git push -q 2>> "$LOG"; then
                    echo "$(date '+%H:%M:%S') ✅ Успех: $MSG" >> "$LOG"
                    osascript -e "display notification \"$MSG. Streamlit Cloud обновится через 30 сек.\" with title \"✅ ROLLS KRAN обновлён\" sound name \"Glass\"" 2>/dev/null || true
                else
                    echo "$(date '+%H:%M:%S') ❌ Пуш не удался. Проверь ~/rolls_watcher.log" >> "$LOG"
                    osascript -e 'display notification "Пуш не удался. Проверь ~/rolls_watcher.log" with title "❌ ROLLS KRAN — ошибка пуша" sound name "Basso"' 2>/dev/null || true
                fi
            fi
        fi

        # Обновляем состояние
        echo "$LATEST_MTIME" > "$STATE"
        LAST_MTIME=$LATEST_MTIME
    fi

    sleep 3
done
EOF
chmod +x "$WATCHER_SCRIPT"
echo "✅ Создан watcher-скрипт: $WATCHER_SCRIPT"

# --- 3. Регистрируем в launchd для автозапуска ---
mkdir -p "$HOME/Library/LaunchAgents"

cat > "$PLIST_FILE" << EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.rolls.watcher</string>
    <key>ProgramArguments</key>
    <array>
        <string>/bin/bash</string>
        <string>$WATCHER_SCRIPT</string>
    </array>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>StandardOutPath</key>
    <string>$LOG_FILE</string>
    <key>StandardErrorPath</key>
    <string>$LOG_FILE</string>
</dict>
</plist>
EOF
echo "✅ Создан plist для автозапуска: $PLIST_FILE"

# Останавливаем старую версию если запущена
launchctl unload "$PLIST_FILE" 2>/dev/null || true

# Запускаем новую
launchctl load "$PLIST_FILE"
echo "✅ Watcher запущен в фоне"

echo ""
echo "🎉 Автоматическая заливка настроена!"
echo ""
echo "📖 Как это работает:"
echo "   �� Watcher работает в фоне 24/7 (даже после перезагрузки Mac)"
echo "   • Каждые 3 секунды проверяет ~/Downloads"
echo "   • Как только появляется новый lks-kp-app*.zip — сам:"
echo "       1. Распаковывает"
echo "       2. Синхронизирует с ~/rolls-kran"
echo "       3. git commit + git push в GitHub"
echo "       4. Показывает уведомление macOS"
echo ""
echo "🔍 Проверка:"
echo "   • Лог:              cat ~/rolls_watcher.log"
echo "   • Статус:           launchctl list | grep rolls"
echo "   • Остановить:       bash ~/Downloads/lks-kp-app/install_auto_watcher.sh --stop"
echo "   • Запустить снова:  bash ~/Downloads/lks-kp-app/install_auto_watcher.sh"
echo ""
echo "✅ Готово. Теперь можешь скачивать мои zip — они будут автоматически коммитить."
