#!/usr/bin/env bash

cd '/home/tmluser/DriveD/Projects/1.python/daily-task-tracker' || exit 1

PID_FILE=".streamlit.pid"

if [ ! -f "$PID_FILE" ]; then
    echo "Task Tracker is not running (PID file not found)."
    exit 0
fi

PID=$(cat "$PID_FILE")

if kill -0 "$PID" 2>/dev/null; then
    kill "$PID"
    echo "Daily Task Tracker stopped."
else
    echo "Task Tracker process is already stopped."
fi

rm -f "$PID_FILE"
