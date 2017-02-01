#!/bin/bash

cd $( cd $(dirname $0) ; pwd )
./mem_limit.sh 500M perl ./send_mail.pl
