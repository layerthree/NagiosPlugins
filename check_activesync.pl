#!/usr/bin/perl
use Time::HiRes qw(gettimeofday tv_interval);
use IO::CaptureOutput qw/capture_exec/;
use Getopt::Std;
use Switch;

# Command arguments
my %options=();
getopts("H:u:p:w:c:", \%options);

# Help message stuff
$help_info = <<END;
    *$0 v1.0*

    Checks Exchange ActiveSync via EWS (Exchange Web Services)
    
    Usage: 
        -H    Hostname of EWS Server (required)
        -u    Username of mailbox to check in username\@domain.local format (required)
        -p    Password for mailbox (required)
        -w    Warning threshold in seconds
        -c    Critical threshold in seconds

    Example:
        $0 -H mail.ews.server -u mailbox -p password [-w 5] [-c 10]

END

# Contants for nagios exit codes
$OKAY = 0;
$WARNING = 1;
$CRITICAL = 2;
$UNKNOWN = 3;

# If we don't have the needed command line arguments, just exit with UNKNOWN.
if(!defined $options{H} || !defined $options{u} || !defined $options{p}){
    print "$help_info Not all required options were specified.\n\n";
    exit $UNKNOWN;
}

# Command variables
$owa_server = $options{H}; #required
$username = $options{u}; #required
$password = $options{p}; #required
$warn = $options{w};
$crit = $options{c};
$ews_path = "/ews/exchange.asmx";

$result = $UNKNOWN; # Default exit code 
$message = "Status UNKNOWN"; # String output to nagios

# XML data submitted to the Exchange Web Service that will give the command to query the Inbox of the mailbox we are checking
$xml_request = '\'<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"><soap:Body><GetFolder xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"><FolderShape><t:BaseShape>Default</t:BaseShape></FolderShape><FolderIds><t:DistinguishedFolderId Id="inbox"/></FolderIds></GetFolder></soap:Body></soap:Envelope>\'';

# Curl command that submits the XML request data to Exchange
$curl = "curl -v -k -u $username:$password --ntlm -L https://$owa_server"."$ews_path -d $xml_request -H \"Content-Type:text/xml\"";

$t0 = [gettimeofday]; # start the clock to time check

# Capture all output - $stdout has the XML result returned by exchange server, $stderr has the http header data (for error control)
my ($stdout, $stderr, $success, $exit_code) = capture_exec( $curl );

$check_time = tv_interval($t0); # $check_time will have how long it took to check the mailbox

$perfdata = "|'Response Time'=$check_time"."s;;;";

# If the connection successful, parse the XML and see if check is also successful. 
# If it isn't, then parse the http header and tell us why

if ($stdout =~ m/<\?xml version=\"1.0\" encoding=\"utf-8\"\?>/){
    switch($stdout){
        case m/<m\:ResponseCode>NoError<\/m\:ResponseCode>/{
            if(($warn && $check_time >= $warn && $check_time < $crit) || ($warn && $check_time >= $warn)){
                $message = "WARNING: Inbox for '$username' polled in $check_time seconds$perfdata\n";
                $result = $WARNING;
            }elsif($crit && $check_time >= $crit){
                $message = "CRITICAL: Inbox for '$username' polled in $check_time seconds$perfdata\n";
                $result = $CRITICAL;
            }else{
                $message = "OK: Inbox for '$username' polled in $check_time seconds$perfdata\n";
                $result = $OKAY;
            }
        }
        case m/<m\:ResponseCode>ErrorMissingEmailAddress<\/m\:ResponseCode>/ {
            $message = "CRITICAL: No mailbox exists for account '$username'\n";
            $result = $CRITICAL;   
        }
        else{
            $message = "CRITICAL: unknown error accessing inbox of '$username'\n";
            $result = $CRITICAL;   
        }
    }
}else{
    switch($stderr){
        case m/HTTP\/1.1 401 Unauthorized/ {
            $message = "CRITICAL: Access denied for username '$username'\n";
            $result = $CRITICAL;
        }
        case m/Could not resolve host/ {
            $message = "CRITICAL: could not resolve host '$owa_server'\n";
            $result = $CRITICAL;
        }
        case m/HTTP\/1.1 404 Not Found/ {
            $message = "CRITICAL: HTTP 404 Error - EWS not found at 'https://$owa_server"."$ews_path'\n";
            $result = $CRITICAL;
        }
        case m/HTTP\/1.1 403 Forbidden/ {
            $message = "CRITICAL: HTTP 403 Error - forbidden at 'https://$owa_server"."$ews_path'\n";
            $result = $CRITICAL ;           
        }
        case m/Operation timed out/ {
            $message = "CRITICAL: Couldn't connect to host '$owa_server' - Operation timed out\n";
            $result = $CRITICAL;
        }
        else{
            $message = "CRITICAL: unknown error connecting to EWS at 'https://$owa_server"."$ews_path'\n";
            $result = $CRITICAL;
        }
    }
}

# exit with message and result
#print STDOUT $message;
print "$stderr\n";
exit $result;