#!/bin/bash

cd /root/sql_job/mailq
./mem_limit.sh 2800M ./send_mail.pl
