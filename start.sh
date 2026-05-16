#!/bin/bash
# Ma'lumotlar papkasini yaratish
mkdir -p /data

# Agar /data/quizbot.db mavjud bo'lmasa, boshlang'ich bazani ko'chiramiz
if [ ! -f /data/quizbot.db ]; then
    echo "Boshlang'ich baza ko'chirilmoqda..."
    cp /app/quizbot.db /data/quizbot.db
    echo "Baza muvaffaqiyatli ko'chirildi!"
else
    echo "Mavjud baza topildi, davom etilmoqda..."
fi

# Botni ishga tushirish
export DATA_DIR=/data
python3 bot.py
