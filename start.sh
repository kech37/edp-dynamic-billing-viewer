#!/bin/sh
# Starts the Energy Consumption Viewer at http://localhost:8123/app/
cd "$(dirname "$0")"
(sleep 1; open "http://localhost:8123/app/") &
node serve.mjs
