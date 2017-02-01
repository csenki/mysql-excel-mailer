#!/usr/bin/perl
use strict;
use MIME::Lite;
use JSON::XS;
use DBI;
use Try::Tiny;
use Excel::Writer::XLSX;
use Encode;
use MIME::Base64;
use utf8;
use File::Basename;
use File::Copy;
#use Data::Dumper;
require "autofit.pm";
require "send_mail_dev_defaults.pm";

##Global declarations
my $xlsfilename , $main::dbh ,$main::decoded_json,$main::line,$main::thrcnt;

############################
#latin1->latin2 -utf convert
############################
sub fix_chars {
    my ($s) = @_;

    $s =~ s/\û/ű/g;
    $s =~ s/õ/ő/g;
    $s =~ s/Õ/Ő/g;
    $s =~ s/Û/Ű/g;

    return $s;
}


#Execel Columname to number
#A colum is the 0
sub EC {
 
  my $name = shift;
 
  if ($name !~ /^[A-Z]+$/i) {
      die "No valid columname: $name\n";
  }
 
  if ($name =~ /^([A-Z])([A-Z]+)$/i) {
      return 26**length($2) * (EC($1)+1) + EC($2);
  }
 
  return ord(uc $name) - 64 - 1;
}

#################################
#Send simple mail with text body
#################################
sub my_sendmail {
    my $from_address = $_[0];
    my $to_address   = $_[1];
    my $subject      = $_[2];
    my $mime_type    = 'text/html';
    

    # Create the initial text of the message
    
    if ( $send_mail_defaults::SMTP_HOST eq "" ) {
        $send_mail_defaults::SMTP_HOST = "127.0.0.1";
    }
    my $message_body = "$_[3]\n";

    my $mime_msg = MIME::Lite->new(
        From    => $from_address,
        To      => $to_address,
        Bcc     => "",
		Subject => "=?UTF-8?B?"
          . encode_base64( encode( "utf8", $subject ), "" ) . "?=",
        
        Type    => $mime_type,
        Data    => $message_body

    ) or die "Error creating MIME body: $!\n";


    MIME::Lite->send( 'smtp', $send_mail_defaults::SMTP_HOST )
      or die "Error sending message: $!\n";

    $mime_msg->send() || print "$send_mail_defaults::SMTP_HOST  Error sending message: $!\n";

}


#################################
###Send mail with attachment
#################################
sub send_xls_to_mail {

    my $from_address = $_[0];
    my $to_address   = $_[1];
    my $subject      = $_[2];
    my $myatach      = $_[4];
    my $mime_type    = 'text/html';
    # Create the initial text of the message

    if ( $send_mail_defaults::SMTP_HOST eq "" ) {
        $send_mail_defaults::SMTP_HOST = "127.0.0.1";
    }
    my $message_body = "$_[3]\n";

    my $mime_msg = MIME::Lite->new(
        From    => $from_address,
        To      => $to_address,
        Bcc     => "",
        Subject => "=?UTF-8?B?"
          . encode_base64( encode( "utf8", $subject ), "" ) . "?=",

        Type    => $mime_type,
        Data    => $message_body

    ) or die "Error creating MIME body: $!\n";
    $mime_msg->attach(
        Type => 'application/octet-stream',
        Path => $myatach,
		Filename =>basename(  $$main::line{"file_name"})
    ) or die "Error attaching file: $!\n";

    
    MIME::Lite->send( 'smtp',$send_mail_defaults::SMTP_HOST )
      or die "Error sending message: $!\n";

    $mime_msg->send() || print "Error sending message: $!\n";
}

######################
#query_to_xlsx
######################
sub query_to_xlsx {
    my $workbook;
    if ( $$main::line{"mode"} eq "2" ) {
        $xlsfilename = "$send_mail_defaults::tempdir/" . basename( $$main::line{"file_name"} . rand(100) );
    }
    if ( $$main::line{"mode"} eq "3" ) {
        $xlsfilename = "$send_mail_defaults::tempdir/" . basename( $$main::line{"file_name"} . rand(100) );
    }
    if ( $$main::line{"mode"} eq "4" ) {
        $xlsfilename = "$send_mail_defaults::tempdir/" . basename( $$main::line{"file_name"} . rand(100) );
    }
    if ( $$main::line{"mode"} eq "5" ) {
        $xlsfilename = "$send_mail_defaults::tempdir/" . basename( $$main::line{"file_name"} . rand(100) );
    }

    $workbook = Excel::Writer::XLSX->new($xlsfilename);
    $workbook->set_properties(
	 author   => $send_mail_defaults::AUTHOR,
    );
    my $default_format = $workbook->add_format( num_format => '@' );
	my $nformat = $workbook->add_format(num_format => '### ### ### ###');
	$nformat->set_border(0);
    $nformat->set_font($send_mail_defaults::FONT_NAME);
    $nformat->set_size($send_mail_defaults::FONT_SIZE);
	
	 
    $default_format->set_font($send_mail_defaults::FONT_NAME);
    $default_format->set_size($send_mail_defaults::FONT_SIZE);
    $default_format->set_border(0);
    if( $main::decoded_json->{"text_wrap"} eq "1"  ){
        $default_format->set_text_wrap();
    }
    if( $main::decoded_json->{"set_optimization"} eq "1"  ){
        $workbook->set_optimization();
    }
    my $bold_format = $workbook->add_format();
    $bold_format->set_font($send_mail_defaults::FONT_NAME);
    $bold_format->set_bold();
    $bold_format->set_size($send_mail_defaults::FONT_SIZE);
    $bold_format->set_border(0);

    my $more_results;
    my $count = 0;
    my $sth   = $main::dbh->prepare(qq($$main::line{"sql_command"}) );
	
#    print $$main::line{"sql_command"};
    $sth->execute || die DBI::err . ": " . $DBI::errstr ;
	
#    print Dumper($sth);
    #trying process header
    do {
        $count++;
	    my @cols;
		if ($sth->{NUM_OF_FIELDS}>0) {
			try{
				@cols = @{ $sth->{NAME} };
			}
			catch{
	        next;
			};
		};
        #Skip empty results
        if ( scalar @cols > 0 ) {

            my $sheet;
			my $rcnt;
            #Create Sheet
            if ( $main::decoded_json->{ "ws" . $count }{'name'} ne "" ) {
                $sheet =
                  $workbook->add_worksheet(
                    $main::decoded_json->{ "ws" . $count }{'name'} );
            }
            else {
                $sheet = $workbook->add_worksheet( "Munkafüzet" . $count );
            }
			my $hun_corr=$send_mail_defaults::DEF_HUN_CORR;
			if ( $main::decoded_json->{ "ws" . $count }{'hun_corr'} ne "" ) {
				$hun_corr=$main::decoded_json->{ "ws" . $count }{'hun_corr'};
            }
            

            if ( $main::decoded_json->{ "ws" . $count }{'autofit'} ne "0" ) {
				$sheet->add_write_handler( qr[\w], \&store_string_widths );
            }			
            
            #Write xlsx header
            my $cnum = 0;
            my $i = 0;
	    if ( not ($main::decoded_json->{ "ws" . $count }{'noheader'} eq "1" )) {
                foreach (@cols) {
				utf8::decode($_);
	            if ($hun_corr eq 1) {
				  $sheet->write( 0, $cnum, fix_chars($_), $bold_format );
				}
				else
				{				
				$sheet->write( 0, $cnum,  $_, $bold_format );
				}
    	            $cnum++;
        	}
            $i = 1;
	    }


            my $row;
			#sorok szama
			$rcnt=$sth->rows;
			#write data to xlsx
			
			if ( $sth->rows > 0 ) {
                while ( $row = $sth->fetchrow_arrayref ) {
				   if ($hun_corr eq 1) {
                    $sheet->write( $i, $_, fix_chars( $row->[$_] ),
                        $default_format )
                      for ( 0 .. $#$row );}
					  else
					  { 		    
					    $sheet->write( $i,  $_,  $row->[$_] ,
                        $default_format )
                      for ( 0 .. $#$row );
					  }
                    $i++;
                }
            }


            if ( $main::decoded_json->{ "ws" . $count }{'autofit'} ne "0" ) {
                autofit_columns($sheet);
            }
            if ( $main::decoded_json->{ "ws" . $count }{'active'} eq "1" ) {
		          $sheet->activate();
            }			
            if ( $main::decoded_json->{ "ws" . $count }{'freeze_panes'} eq "1" ) {
		          $sheet->freeze_panes( 1, 0 ); ;
            }			

            $cnum = 0;
			#Column length from json
            foreach (@cols) {
                if ( $main::decoded_json->{ "ws" . $count }{ "col_len" . $cnum } ne
                    "" )
                {
                    $sheet->set_column( $cnum, $cnum,
                        $main::decoded_json->{ "ws" . $count }{ "col_len" . $cnum } );
                }
                $cnum++;
            }
		my $rcnt2=$rcnt+1;
		if ($main::decoded_json->{ "ws" . $count }{'tonum'} ne "" ) {
		    #example tonum  A1:A*
			#example tonum  A1:A100
			#example tonum:  A2:A*,J2:J*
			my $tonum=$main::decoded_json->{ "ws" . $count }{'tonum'};
			$tonum=~ s/\*/$rcnt2/g;
			$sheet->conditional_formatting($tonum,{type => 'no_errors', format => $nformat,});
		};
        if ( $main::decoded_json->{ "ws" . $count }{'autofilter'} eq "1" ) {
		          $sheet->autofilter(0,0,$rcnt,$cnum-1);
        };
        
		}

#    } while ( $sth->more_results || die DBI::err . ": " . $DBI::errstr . $_  );
    } while ( $sth->more_results  );

    $workbook->close;
    if (( $$main::line{"mode"} eq "3" ) || ( $$main::line{"mode"} eq "4" ) || ( $$main::line{"mode"} eq "5" )  ) {

	my $tmpfile;

	$tmpfile="$send_mail_defaults::basedir" . $$main::line{"file_name"};
	#Make "safe" file names
	$tmpfile=~ s/\.\.//g;
        move($xlsfilename,$tmpfile);
    }
}

#####
#query_to_xlsx_caller
#####
sub query_to_xlsx_caller{
my $stderr;
		{
		  local *STDERR;
          open STDERR, ">", \$stderr;
		  query_to_xlsx; 
		}
		if ($stderr ne "") {
		 die  $stderr;
		}
}
###################
#Send select result 
###################
sub select_to_mail {
    try {
        if ( $$main::line{"xls_opts"} ne "" ) {
			$main::decoded_json=JSON::XS->new->utf8->decode(encode("UTF8",$$main::line{"xls_opts"}));
        }
        query_to_xlsx_caller;
        send_xls_to_mail(
            $$main::line{"msg_from"}, $$main::line{"msg_to"}, $$main::line{"msg_subject"},
            $$main::line{"msg_body"}, $xlsfilename
        );
        unlink $xlsfilename;
    }
    catch {
        my_sendmail( $$main::line{"msg_from"}, $$main::line{"on_error"},
            "ERROR: " . $$main::line{"msg_subject"}, $_ );
    };
}

######################
#Save select to file
######################
sub select_to_file {
    try {
        if ( $$main::line{"xls_opts"} ne "" ) {
			$main::decoded_json=JSON::XS->new->utf8->decode(encode("UTF8",$$main::line{"xls_opts"}));

        }
        query_to_xlsx_caller;		
        my_sendmail(
            $$main::line{"msg_from"},    $$main::line{"msg_to"},
            $$main::line{"msg_subject"}, $$main::line{"msg_body"}
        );
    }
    catch {
        my_sendmail( $$main::line{"msg_from"}, $$main::line{"msg_to"},
            "ERROR: " . $$main::line{"msg_subject"}, $_ );
    };
}


######################
#Save select to file, and send mail if error
######################
sub select_to_file_if_error {
    try {
        if ( $$main::line{"xls_opts"} ne "" ) {
			$main::decoded_json=JSON::XS->new->utf8->decode(encode("UTF8",$$main::line{"xls_opts"}));

        }
        query_to_xlsx_caller;
    }
    catch {
        my_sendmail( $$main::line{"msg_from"}, $$main::line{"msg_to"},
            "ERROR: " . $$main::line{"msg_subject"}, $_ );
    };
}


#######################################
#Save select to file and send in email
#######################################
sub select_to_file_and_send {
    try {
        if ( $$main::line{"xls_opts"} ne "" ) {
			$main::decoded_json=JSON::XS->new->utf8->decode(encode("UTF8",$$main::line{"xls_opts"}));

        }
     
		query_to_xlsx_caller;
		

        send_xls_to_mail(
            $$main::line{"msg_from"}, $$main::line{"msg_to"}, $$main::line{"msg_subject"},
            $$main::line{"msg_body"}, "$send_mail_defaults::basedir" . $$main::line{"file_name"}
        );
    }
    catch {
        my_sendmail( $$main::line{"msg_from"}, $$main::line{"on_error"},
            "ERROR: " . $$main::line{"msg_subject"}, $_ );
    };
}

#########################################################################
#MAIN
#########################################################################


$main::thrcnt=`ps ax |grep -c send_mail.pl`;
if ( $main::thrcnt >15 ) {
# print $main::thrcnt;
 exit 0;
}
$main::dbh = DBI->connect(
"DBI:mysql:database=$send_mail_defaults::SQL_DB;host=$send_mail_defaults::SQL_HOST;".
"user=$send_mail_defaults::SQL_USER;password=$send_mail_defaults::SQL_PASS;mysql_multi_statements=1;".
"{RaiseError => 1, PrintError => 1}"
) or die "Couldn't connect to database: " . "$DBI::errstr";

$main::dbh->{'mysql_enable_utf8'} = 1;
$main::dbh->do('SET NAMES utf8');
my $sql =
qq(Select rowid, msg_from,msg_to,msg_subject,msg_body,mode,sql_command,on_error,file_name,trim(xls_opts) as xls_opts 
from $send_mail_defaults::SQL_TABLE as t where msg_ready=0 limit 1);


my $sth = $main::dbh->prepare($sql);
$sth->execute();

while ( $main::line = $sth->fetchrow_hashref ) {
    try
    {
     $main::dbh->do( "update $send_mail_defaults::SQL_TABLE set msg_ready=2 where rowid=".$$main::line{"rowid"} );

        if ( $$main::line{"mode"} eq "1" ) {
	    #sima szoveges ertesites
            my_sendmail(
                $$main::line{"msg_from"}, $$main::line{"msg_to"},
                $$main::line{"msg_subject"}, $$main::line{"msg_body"}
            );
        }
        elsif ( $$main::line{"mode"} eq "2" ) {
	    #select ->xlslx->email
            select_to_mail;
        }
        elsif ( $$main::line{"mode"} eq "3" ) {
	    #select->xlsx->file
            select_to_file;

        }
        elsif ( $$main::line{"mode"} eq "4" ) {
	    #select->xlsx->file->email
            select_to_file_and_send;
	}
        elsif ( $$main::line{"mode"} eq "5" ) {
            select_to_file_if_error;

        }		
     $main::dbh->do( "update $send_mail_defaults::SQL_TABLE set msg_ready=1 where rowid="
              . $$main::line{"rowid"} );
    } 
    catch { print "Error: ". $_ ."\n";  };


   $sth->finish;
   $sth->execute();
  }
__END__
