# mysql-excel-mailer
Send email from mysql stored procedure. Email atachment (excel) generated from sql select with multiple resultset.
How to use:

Create table mailq:
cat create.sql | mysql databasename

install dependencies:
apt install libmime-lite-perl libjson-xs-perl libdbi-perl libtry-tiny-perl libexcel-writer-xlsx-perl libmime-base64-urlsafe-perl  libfile-copy-recursive-perl 

edit send_mail_defaults.pm








Examples usage:


  INSERT INTO tools.mailq (msg_from, msg_to, msg_subject, msg_body,  `mode`, sql_command, on_error, file_name, xls_opts)
                VALUES ('FROM@example.com', #from
                'to1@example.com,to2@example.com', #to
                concat('SUBJECT - ',date_format(now(),'%Y-%m-%d')), #msg_subject
                'Dear ..!<br><br>I send it. <br><br>', #msg_body
                2, #mode (Send results in email)
                "select 1 as res1 ; select 2 as res2; select 3 as res3;select 4 as res4;
                ", # sql_command 
                 'error@example.com', #on_error
                 concat('filename_',date_format(now(),'%Y-%m-%d'),'.xlsx'), #file_name
                 '
                {
                                 "ws1": {
                                   "name": "Shhet1",
                                   "autofit": "1",
                                   "hun_corr": "1"
                                 },
                                 "ws2": {
                                   "name": "Sheet2",
                                   "autofit": "1",
                                   "hun_corr": "1"
                                 },
                                 "ws3": {
                                   "name": "Sheet3",
                                   "autofit": "1",
                                   "hun_corr": "1"
                                 },
                                 "ws4": {
                                   "name": "Sheet4",
                                   "autofit": "1",
                                   "hun_corr": "1"
                                 }                             
                 }
                   ' # xls_opts
                 );
