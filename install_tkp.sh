#!/bin/bash
# =============================================================================
# install_tkp.sh — установка команды `tkp` для быстрой заливки zip → GitHub
#
# Как пользоваться:
#   1. Скачай lks-kp-app.zip в ~/Downloads
#   2. Распакуй его один раз (двойной клик на zip)
#   3. Открой Терминал и выполни:
#        bash ~/Downloads/lks-kp-app/install_tkp.sh
#   4. Скрипт спросит у тебя GitHub Personal Access Token (см. инструкцию ниже)
#   5. Готово — теперь после скачивания моих zip достаточно ввести в терминале:
#        tkp
#
# Как получить PAT:
#   1. Открой https://github.com/settings/tokens
#   2. Generate new token (classic) → отметь `repo` → Generate
#   3. Скопируй токен (начинается с ghp_...)
# =============================================================================

set -e

REPO_URL="github.com/a89817778886-sudo/-.git"
REPO_DIR="$HOME/rolls-kran"
GITHUB_USER="a89817778886-sudo"

echo "🚀 Установка tkp для ROLLS KRAN"
echo ""

# --- 1. Клонирование репозитория, если ещё нет ---
if [ -d "$REPO_DIR/.git" ]; then
    echo "✅ Репозиторий уже склонирован: $REPO_DIR"
else
    echo "📥 Первичная настройка — нужен GitHub Personal Access Token"
    echo ""
    echo "   Если у тебя ещё нет токена:"
    echo "   1. Открой https://github.com/settings/tokens"
    echo "   2. Generate new token → classic → отметь галочку 'repo' → Generate"
    echo "   3. Скопируй токен (начинается с ghp_...)"
    echo ""
    read -sp "🔑 Вставь сюда токен (начинается с ghp_): " GH_TOKEN
    echo ""
    if [ -z "$GH_TOKEN" ]; then
        echo "❌ Токен не введён. Отмена."
        exit 1
    fi
    echo "📥 Клонирую репозиторий в $REPO_DIR ..."
    git clone "https://${GITHUB_USER}:${GH_TOKEN}@${REPO_URL}" "$REPO_DIR"
    # Сохраняем токен в macOS keychain, чтобы больше не спрашивать
    cd "$REPO_DIR"
    git config credential.helper osxkeychain 2>/dev/null || true
    echo "✅ Репозиторий склонирован"
fi

# --- 2. Создаём скрипт update_rolls.sh ---
UPDATE_SCRIPT="$HOME/update_rolls.sh"
cat > "$UPDATE_SCRIPT" << 'EOF'
#!/bin/bash
set -e
cd ~/rolls-kran

# Ищем последний zip с моделью имени
ZIP=$(ls -t ~/Downloads/lks-kp-app*.zip 2>/dev/null | head -1)
if [ -z "$ZIP" ]; then
    echo "❌ В ~/Downloads нет файла lks-kp-app*.zip"
    echo "   Скачай мой zip в папку Загрузки и попробуй снова."
    exit 1
fi

echo "📦 Беру zip: $ZIP"
rm -rf /tmp/lks-new
mkdir -p /tmp/lks-new
unzip -qo "$ZIP" -d /tmp/lks-new

# Синхронизируем в репо, исключая локальные файлы
rsync -a --delete \
    --exclude='.git' \
    --exclude='crm.db' \
    --exclude='crm_files' \
    --exclude='crm_backups' \
    --exclude='.dadata_token' \
    --exclude='history_memory.json' \
    --exclude='install_tkp.sh' \
    /tmp/lks-new/lks-kp-app/ ./

git add -A
if git diff --cached --quiet; then
    echo "⚠️  Нет изменений — код уже актуален."
    exit 0
fi

echo ""
echo "→ Что изменилось:"
git status --short | head -20
echo ""

MSG="update $(date '+%d.%m.%Y %H:%M')"
git commit -q -m "$MSG"
git push -q
echo "✅ Готово: $MSG"
echo "   Streamlit Cloud обновится через ~30 секунд."
EOF
chmod +x "$UPDATE_SCRIPT"
echo "✅ Создан скрипт: $UPDATE_SCRIPT"

# --- 3. Добавляем алиас tkp в .zshrc ---
ZSHRC="$HOME/.zshrc"
if grep -q "alias tkp=" "$ZSHRC" 2>/dev/null; then
    echo "✅ Алиас tkp уже настроен в ~/.zshrc"
else
    cat >> "$ZSHRC" << 'EOF'

# ROLLS KRAN — быстрая заливка zip в GitHub
alias tkp='~/update_rolls.sh'
EOF
    echo "✅ Добавлен алиас tkp в ~/.zshrc"
fi

echo ""
echo "🎉 Установка завершена!"
echo ""
echo "📖 Как пользоваться:"
echo "   1. Скачай мой новый zip в ~/Downloads"
echo "   2. В терминале введи одно слово:"
echo ""
echo "        tkp"
echo ""
echo "   3. Готово — код улетит в GitHub, Streamlit Cloud обновится."
echo ""
echo "⚠️  Для активации команды tkp в этой сессии выполни:"
echo ""
echo "        source ~/.zshrc"
echo ""
echo "   Или просто открой новое окно терминала — там tkp уже будет работать."
