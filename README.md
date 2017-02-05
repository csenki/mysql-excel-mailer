# mysql-excel-mailer
Send email from MySQL (or MariaDB) stored procedure. Email attachment (excel - XLSX) generated from SQL SELECT with multiple result set.

How to use:

install dependencies (on UBUNTU):
<code>
apt install libmime-lite-perl libjson-xs-perl libdbi-perl libtry-tiny-perl libexcel-writer-xlsx-perl libmime-base64-urlsafe-perl  libfile-copy-recursive-perl cgmanager cgmanager-utils
</code>

Download:
<pre>
cd /opt
git clone https://github.com/csenki/mysql-excel-mailer
cd mysql-excel-mailer
</pre>

Create table mailq:
<pre>
cat create.sql | mysql databasename
</pre>

<code>
cp send_mail_defaults_template.pm send_mail_defaults.pm
</code>

Edit send_mail_defaults.pm 



Insert a line into mailq (like example)

Start manual send_mail.pl (<code>perl send_mail.pl</code>) or add to /etc/crontab (every 2 minutes start send_mail.pl):

<pre>
*/2 * * * *	root	/opt/mysql-excel-mailer/send_mail_starter.sh
</pre>

For more information see our WIKI pages!

Enyoj it!



Example usage:

<pre>
  INSERT INTO mailq (msg_from, msg_to, msg_subject, msg_body,  `mode`,
                sql_command, on_error, file_name, xls_opts)
                VALUES ('FROM@example.com',                            #from
                'to1@example.com,to2@example.com',                     #to
                concat('SUBJECT - ',date_format(now(),'%Y-%m-%d')),    #msg_subject
                'Dear ..!<br><br>I send it. <br><br>', #msg_body
                2, #mode (Send results in email)
                '
                 select 1 as res1 ; select 2 as res2; select 3 as res3;select 4 as res4;
                ', # sql_command
                 'error@example.com', #on_error
                 concat('filename_',date_format(now(),'%Y-%m-%d'),'.xlsx'), #file_name
                 '
                {
                                 "ws1": {
                                   "name": "Sheet1",
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
                '                                                      # xls_opts
                 );
</pre>
