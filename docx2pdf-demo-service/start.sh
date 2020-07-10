#!/bin/sh
LOGS_DIR=logs

if [ ! -d "${LOGS_DIR}" ]
then
  mkdir "${LOGS_DIR}"
fi

python3 doc2pdf-demo-service.py doc2pdf-demo-service.conf
