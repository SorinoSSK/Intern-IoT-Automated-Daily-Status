#!/bin/sh -e

tmux new-session -d -s statusSession 'python3.11 /home/adv/automated_units_status/unit_status.py'
tmux set-option -t statusSession:0 remain-on-exit
