#!/usr/bin/env bash
python3 ./get_accounts.py
python3 ./investment_calculator.py "$@"
