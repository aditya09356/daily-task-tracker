#!/usr/bin/env bash
cd '/home/tmluser/DriveD/Projects/1.python/daily-task-tracker/' || exit 1

# Avoid starting a duplicate instance
if [ -f .streamlit.pid ] && kill -0 "$(cat .streamlit.pid)" 2>/dev/null; then
    echo "Task Tracker is already running."
    exit 0
fi

nohup python3 -m streamlit run app.py \
    --server.headless=true \
    > streamlit.log 2>&1 &

echo $! > .streamlit.pid
sleep 2
xdg-open http://localhost:8501 >/dev/null 2>&1 &
