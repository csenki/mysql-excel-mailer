#!/bin/bash

cd $( cd $(dirname $0) ; pwd )
./mem_limit.sh 2800M ./send_mail.pl
