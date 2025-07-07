#!/bin/bash

# Скрипт автоматической установки системы перевода документов
# Literary Translator v3.0.1

set -e  # Остановка при ошибке

# Цвета для вывода
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}🚀 Установка Literary Translator v3.0.1${NC}"
echo -e "${BLUE}=======================================${NC}"

# Проверка операционной системы
OS="$(uname -s)"
case "${OS}" in
    Linux*)     MACHINE=Linux;;
    Darwin*)    MACHINE=Mac;;
    CYGWIN*)    MACHINE=Cygwin;;
    MINGW*)     MACHINE=MinGw;;
    *)          MACHINE="UNKNOWN:${OS}"
esac

echo -e "${BLUE}🔍 Обнаружена система: ${MACHINE}${NC}"

# Проверка Python
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}❌ Python 3 не найден. Установите Python 3.8 или новее.${NC}"
    exit 1
fi

PYTHON_VERSION=$(python3 -c 'import sys; print(".".join(map(str, sys.version_info[:2])))')
echo -e "${GREEN}✅ Python ${PYTHON_VERSION} найден${NC}"

# Проверка pip
if ! command -v pip3 &> /dev/null; then
    echo -e "${RED}❌ pip3 не найден. Установите pip3.${NC}"
    exit 1
fi

echo -e "${GREEN}✅ pip3 найден${NC}"

# Создание виртуального окружения
echo -e "${YELLOW}📦 Создание виртуального окружения...${NC}"
if [ ! -d "venv" ]; then
    python3 -m venv venv
    echo -e "${GREEN}✅ Виртуальное окружение создано${NC}"
else
    echo -e "${YELLOW}⚠️  Виртуальное окружение уже существует${NC}"
fi

# Активация виртуального окружения
echo -e "${YELLOW}🔧 Активация виртуального окружения...${NC}"
source venv/bin/activate

# Обновление pip
echo -e "${YELLOW}🔄 Обновление pip...${NC}"
pip install --upgrade pip

# Установка зависимостей
echo -e "${YELLOW}📥 Установка зависимостей...${NC}"
pip install -r requirements.txt

# Проверка .env файла
echo -e "${YELLOW}🔐 Проверка конфигурации...${NC}"
if [ ! -f ".env" ]; then
    if [ -f ".env.example" ]; then
        cp .env.example .env
        echo -e "${YELLOW}⚠️  Создан .env файл из примера${NC}"
        echo -e "${YELLOW}📝 Не забудьте настроить API ключи в .env файле!${NC}"
    else
        echo -e "${RED}❌ Не найден .env.example файл${NC}"
    fi
else
    echo -e "${GREEN}✅ .env файл найден${NC}"
fi

# Создание директории для логов
echo -e "${YELLOW}📁 Создание директории для логов...${NC}"
mkdir -p logs
echo -e "${GREEN}✅ Директория logs создана${NC}"

# Тестирование установки
echo -e "${YELLOW}🧪 Тестирование установки...${NC}"
if python3 -c "import docx, PIL, requests; print('Все модули импортированы успешно')" 2>/dev/null; then
    echo -e "${GREEN}✅ Все зависимости установлены корректно${NC}"
else
    echo -e "${RED}❌ Ошибка при импорте зависимостей${NC}"
    exit 1
fi

# Финальные инструкции
echo -e "${BLUE}🎉 Установка завершена!${NC}"
echo -e "${BLUE}===================${NC}"
echo ""
echo -e "${GREEN}Для использования системы:${NC}"
echo -e "1. Активируйте виртуальное окружение:"
echo -e "   ${YELLOW}source venv/bin/activate${NC}"
echo ""
echo -e "2. Настройте API ключи в .env файле:"
echo -e "   ${YELLOW}nano .env${NC}"
echo ""
echo -e "3. Запустите перевод:"
echo -e "   ${YELLOW}python3 literary_translate.py input.docx output.docx${NC}"
echo ""
echo -e "4. Для получения справки:"
echo -e "   ${YELLOW}python3 literary_translate.py --help${NC}"
echo ""
echo -e "${BLUE}📚 Документация: README.md${NC}"
echo -e "${BLUE}🐛 Проблемы: https://github.com/your-repo/literary-translator/issues${NC}" 