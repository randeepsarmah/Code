# Data Analysis  of ERS 8K and create input file to the Config Generator Script
#
# Author  : Randeep Sarmah
# Revision: 1.2
# Release : Jan 15th, 2018
#
# Purpose:
#
#   This Perl Script will help the Extreme engineers to input a ERS 8K data "show run" , "show tech" , "show config" (both ppcli and nncli ) and create a Excel file for 
#	discussion with the Customer and after this Excel Checkpoint file is finalized. We can input this excel file to config genrator script to generate configuration for each switch
#
#   
#
#  
#
# Installation Instructions:
#
#   1. Need Perl to run this script  and few modules such as :
#				
# 	use List::MoreUtils qw(uniq);
#	use Excel::Writer::XLSX;
# 
# Usage Instructions:
#
# 	1.  perl dataanalysis.pl Data_for_SW121_DNGS.txt

#
#
# Known Issue/Limitation :
#
#	1.  Only works on ERS 8K , may work on ERS 8300 but need enhacement for ERS 8300.
#	2.  Doesn't work on Stackable or ERS edge switches 
#
#
# Planned Feature Enhacement  :
#
#	1.  May incorporate ERS 8300 , if more data or switch replacement is done 
#	
#		 
#
#    
#   NOTE : This Script is created for Extreme Internal use only and shouldn't be share with any external Customers
#	
#    UPDATE : 
#
#  		# Version 1.1 ..Change the colour from Avaya Red to Extreme Purple
#  		# Version 1.2 ..Parsing of MLT of ERS8800/ERS8300
#		# Version 1.3 ...Pasrsing of EXOS with few limitation on port SFP type field . 
#		# Version 1.5 ...Pasrsing of EOS 


#use strict;
#use warnings;
use Excel::Writer::XLSX;
use List::MoreUtils qw(uniq);
use Data::Dumper;



my $firstfile = $ARGV[0];           # store the 1st argument into the variable
open my $FH1, '<', $firstfile or die $!; # open the file using lexically scoped filehandle

our $rowcounter = 0;

$count = 1;

	$time = strftime("%Y-%m-%d %H.%M_", localtime);
	$excelName = join ".", $hostname, "xlsx";


  # Create a new Excel workbook
    my $workbook = Excel::Writer::XLSX->new( $time.$excelName );
	die "Problems creating new Excel file: $!" unless defined $workbook;

    # Add a worksheet
    my   $worksheet = $workbook->add_worksheet('Data Analysis');
	
	#Calibri 10 and Wordwrap
	my $format_def = $workbook->add_format();
     $format_def->set_font('Calibri');
	 $format_def->set_size('9');
     $format_def->set_text_wrap();
     $format_def->set_align( 'left' );
	 
	#Purple for Extreme 
	my $format_purple = $workbook->add_format();
	   $format_purple->set_bold();
	   $format_purple->set_color( 'white' );
	   $format_purple->set_bg_color( '#9400D3' );
	   $format_purple->set_align( 'center' );
	   $format_purple->set_font('Calibri');
	   $format_purple->set_size('10');
	  $format_purple->set_text_wrap();
	   $format_purple->set_border( '1' );
	   
	  	#Yellow for New  Extreme 
	my $format_yellow = $workbook->add_format();
	   $format_yellow->set_bold();
	   $format_yellow->set_color( 'black' );
	   $format_yellow->set_bg_color( '#FFFF33' );
	   $format_yellow->set_align( 'center' );
	   $format_yellow->set_font('Calibri');
	   $format_yellow->set_size('10');
	   $format_yellow->set_text_wrap();
	   $format_yellow->set_border( '1' );
	
	 
	 # Set Coloum lengh
	$worksheet->set_column( 0, 0, 15);  
	$worksheet->set_column( 1, 1, 8 );
    $worksheet->set_column( 2, 2, 20 );   # Columns E-F width set to 5
    $worksheet->set_column( 3, 3, 8 );   # Column  F-H   width set to 12
	$worksheet->set_column( 4, 4, 15 );
	$worksheet->set_column( 5, 5, 12);
	$worksheet->set_column( 6, 6, 12 );   # Column  I   width set to 12
    $worksheet->set_column( 7, 7, 6 );   # Columns J-K width set to 8
	$worksheet->set_column( 8, 8, 12 );   # Columns J-K width set to 8
	$worksheet->set_column( 9, 9, 10 );   # Columns J-K width set to 8
	$worksheet->set_column( 10, 10, 15 );
	$worksheet->set_column( 11, 11, 6 );
	$worksheet->set_column( 12, 12, 6 );
	$worksheet->set_column( 13, 13, 15 );
	$worksheet->set_column( 14, 14, 8 );
	$worksheet->set_column( 15, 15, 8 );
	$worksheet->set_column( 16, 16, 10 );
	
	$worksheet->set_column( 17, 17, 6 );
	$worksheet->set_column( 18, 18, 8 );
	$worksheet->set_column( 19, 19, 10 );
	$worksheet->set_column( 20, 20, 30 );
	$worksheet->set_column( 21, 21, 25 );
	$worksheet->set_column( 22, 22, 25 );
	$worksheet->set_column( 23, 23, 9);
	$worksheet->set_column( 24, 25, 8);
	$worksheet->set_column( 26, 26, 10);
	 
	 
	 #### Setting few value ....as disable , which is default state. 

	foreach $mltID (1..512) {$MLT_DATA{"MLT#$mltID#MLTTYPE"}= "MLT";}
	 
	 
	 


############## This part of opening the file for EXOS is to just make a Perl HASH for the Port No to the Description field. 

while(<$FH1>)
{
	#### Port to the Name/Description to port No mapping 
	if ($_=~/configure\s+ports\s+(.*?)\s+display-string\s+(.*?)\s*$/){
	
		chomp ();
		my $portID = $1;
		my $portDesc = $2;
		
				if (( $PORT__DATA_NAME{"$portDesc#PORTID"}) eq "") {     ###  Checking if we already have port ID assigned to the name , if not just push the ID
				
					push (@portNames, $portDesc );      #############  Creating another HASH based on vlan Name to VLAN ID , as EXOS reference everything via VLAN NAME
				
					$PORT__DATA_NAME{"$portDesc#PORTID"} = $portID;		### May be have to put a logic ..if the value not already present push it , else don;t do it. 
					
				} else {                                                 ### Else if value is already present , append it with a  comma 
				
				$portDesc ="DUPLICATE-$count-$portDesc";
				$count ++;
				
				push (@portNames, $portDesc );      #############  Creating another HASH based on vlan Name to VLAN ID , as EXOS reference everything via VLAN NAME
				$PORT__DATA_NAME{"$portDesc#PORTID"} = "$portID";   ###########  Creating a new HASH, does if we have duplicate name rather than replacing it we are putting "," and append
				
				}
	}
	
	
	if ($_=~/(.*?)\s+(\d+)\s+(.*)\s+ANY\s+(\d+ \/\d+)\s+(.*)/){
	
		my $vlanName = $1;
		$vlanName = trim($vlanName);
		my $vlanID = $2;
		$vlanID = trim($vlanID);
		
		push(@vlanIDs,$vlanID);
		$VLAN_DATA{"$vlanID#VLANNAME#VLAN"}=$vlanName;
		
		push (@vlanNames, $vlanName );      #############  Creating another HASH based on vlan Name to VLAN ID , as EXOS reference everything via VLAN NAME
		$VLAN__DATA_NAME{"$vlanName#VLANID"}=$vlanID;   ###########  Creating a new HASH
			
	}
	
	###  armourctr       4019 10.0.0.101     /30  -f-----mop-------------- ANY    1 /1   VR-Default 
	### cgr-cityhall2x  4088 -------------------------------------------- ipx_8022 1 /1   VR-Default 
	### servermgmt_60   2960 -------------------------------------------- ANY    1 /1   VR-Default
	
	if ($_=~/(.*?)\s+(\d+)\s+(\d+.\d+.\d+.\d+)\s+(\/(\d+))\s+(.*?)\s+(ANY|IP|ipx_8022)\s+((\d+)\s+\/(\d+))\s+(.*)/){
	
	
		my $vlanName = $1;
		$vlanName = trim($vlanName);
		my $vlanID = $2;
		$vlanID = trim($vlanID);
		
		push(@vlanIDs,$vlanID);
		$VLAN_DATA{"$vlanID#VLANNAME#VLAN"}=$vlanName;
		
		push (@vlanNames, $vlanName );      #############  Creating another HASH based on vlan Name to VLAN ID , as EXOS reference everything via VLAN NAME
		$VLAN__DATA_NAME{"$vlanName#VLANID"}=$vlanID;   ###########  Creating a new HASH
			
	}
	
	
	if ($_=~/(.*?)\s+(\d+)\s+(\d+.\d+.\d+.\d+)\s+(\/(\d+))\s+(.*?)\s+(ANY|IP|ipx_8022)\s+((\d+)\/(\d+))\s+(.*)/){
	
	
		my $vlanName = $1;
		$vlanName = trim($vlanName);
		my $vlanID = $2;
		$vlanID = trim($vlanID);
		
		push(@vlanIDs,$vlanID);
		$VLAN_DATA{"$vlanID#VLANNAME#VLAN"}=$vlanName;
		
		push (@vlanNames, $vlanName );      #############  Creating another HASH based on vlan Name to VLAN ID , as EXOS reference everything via VLAN NAME
		$VLAN__DATA_NAME{"$vlanName#VLANID"}=$vlanID;   ###########  Creating a new HASH
	
	}
	
	if ($_=~/(.*?)\s+(\d+)(\s+[A-Za-z\-]{44}\s+)(ANY|IP|ipx_8022)\s+((\d+)\s+\/(\d+))\s+(.*)/){
	
	
		my $vlanName = $1;
		$vlanName = trim($vlanName);
		my $vlanID = $2;
		$vlanID = trim($vlanID);
		
		push(@vlanIDs,$vlanID);
		$VLAN_DATA{"$vlanID#VLANNAME#VLAN"}=$vlanName;
		
		push (@vlanNames, $vlanName );      #############  Creating another HASH based on vlan Name to VLAN ID , as EXOS reference everything via VLAN NAME
		$VLAN__DATA_NAME{"$vlanName#VLANID"}=$vlanID;   ###########  Creating a new HASH
	
	}
	if ($_=~/(.*?)\s+(\d+)(\s+[A-Za-z\-]{44}\s+)(ANY|IP|ipx_8022)\s+((\d+)\/(\d+))\s+(.*)/){
	
	
		my $vlanName = $1;
		$vlanName = trim($vlanName);
		my $vlanID = $2;
		$vlanID = trim($vlanID);
		
		push(@vlanIDs,$vlanID);
		$VLAN_DATA{"$vlanID#VLANNAME#VLAN"}=$vlanName;
		
		push (@vlanNames, $vlanName );      #############  Creating another HASH based on vlan Name to VLAN ID , as EXOS reference everything via VLAN NAME
		$VLAN__DATA_NAME{"$vlanName#VLANID"}=$vlanID;   ###########  Creating a new HASH
	
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
}

open my $FH1, '<', $firstfile or die $!; # open the file using lexically scoped filehandle

while(<$FH1>)
{
	chomp();

	#vlan 1 ip create 142.49.150.1/255.255.0.0 mac_offset 0
	
	if ($_=~/vlan\s+1\s+(ip)\s+create\s+(\d+.\d+.\d+.\d+)\/(\d+.\d+.\d+.\d+)\s+(mac_offset)\s+(\d+)/){
	
		my $vlanID = "1";
		my $vlanIP = $2;
		my $vlanMask = $3;

			push(@vlanIDs,$vlanID);
					
			$VLAN_DATA{"$vlanID#VLANIP#VLAN"}=$vlanIP;
			$VLAN_DATA{"$vlanID#VLANMASK#VLAN"}=$vlanMask;
			$VLAN_DATA{"VLAN_BY_IP#$vlanIP"}=$vlanID;
	
	}
	if ($_=~/vlan\s+(\d+)\s+(ip)\s+create\s+(\d+.\d+.\d+.\d+)\/(\d+.\d+.\d+.\d+)\s+(mac_offset)\s+(\d+)/){
	
		my $vlanID = $1;
		my $vlanIP = $3;
		my $vlanMask = $4;

			push(@vlanIDs,$vlanID);
					
			$VLAN_DATA{"$vlanID#VLANIP#VLAN"}=$vlanIP;
			$VLAN_DATA{"$vlanID#VLANMASK#VLAN"}=$vlanMask;
			$VLAN_DATA{"VLAN_BY_IP#$vlanIP"}=$vlanID;
	
	}
	#### vlan 150 ip create 10.200.15.92/255.255.255.224  (No Mac Offset on ERS 8300)
		
	if ($_=~/vlan\s+(\d+)\s+(ip)\s+create\s+(\d+.\d+.\d+.\d+)\/(\d+.\d+.\d+.\d+)\s+/){
	
		my $vlanID = $1;
		my $vlanIP = $3;
		my $vlanMask = $4;

			push(@vlanIDs,$vlanID);
					
			$VLAN_DATA{"$vlanID#VLANIP#VLAN"}=$vlanIP;
			$VLAN_DATA{"$vlanID#VLANMASK#VLAN"}=$vlanMask;
			$VLAN_DATA{"VLAN_BY_IP#$vlanIP"}=$vlanID;
	
	}
	
	
	

	if ($_=~/vlan\s+(\d+)\s+(create)\s+(byport)\s+(\d+)\s+(.*)/){   ## Here I'm extracting the VLAN Name config 
	
			my $vlanID = "$1";
			$VN = $5;				## Storing the VLAN whole section into a variable 
		
			push(@vlanIDs,$vlanID);
		
		 if ($VN =~ /"(.+?)"/) {	## This regex just extract whatever is between the ""
			
			$vlanName = $1; 
			$VLAN_DATA{"$vlanID#VLANNAME#VLAN"}=$vlanName;
						
		}
	
	}
	
	## vlan 150 ip vrrp 150 address 10.200.15.65   (On ERS 8300 , may work on ERS 8600 as well)
	
	if ($_=~/vlan\s+(\d+)\s+ip\s+vrrp\s+(\d+)\s+address\s+(\d+.\d+.\d+.\d+)\s+/){
	
			my $vlanID = $1;
			my $vlanVrrpId = $2;
			my $vlanVrrpIp = $3;
		
			push(@vlanIDs,$vlanID);
		
		#print "*****$vlanID.......$vlanVrrpId.....$vlanVrrpIp\n";
				
			$VLAN_DATA{"$vlanID#VLANVRRPID#VRRPID"} = $vlanVrrpId;
			$VLAN_DATA{"$vlanID#VLANVRRPIP#VRRPIP"} = $vlanVrrpIp;
		
	
	}
	
	
	#ip dhcp-relay create-fwd-path agent 172.22.36.1 server 142.49.128.100 mode bootp_dhcp state enable
	
	if ($_ =~/ip\s+dhcp-relay\s+create-fwd-path\s+agent\s+(\d+.\d+.\d+.\d+)\s+server\s+(\d+.\d+.\d+.\d+)\s+(.*)/){
	
		$vlanIP = $1; 
		$serverIP = $2; 
		$vlanID = $VLAN_DATA{"VLAN_BY_IP#$vlanIP"};

		
		#if(exists($portArray{$interface}{name})){
		$VLAN_DATA{"$vlanID#VLANIP#DHCPSERVER"} .= ",$serverIP";
		
		
		
	
	}
	
	if ($_ =~ /((\d+)\/(\d+))\s+(\d+)\s+(.*?)\s+(true|false)\s+(true|false)\s+(\d+)\s+(.*?)\s+(up|down)\s+(up|down)/) {
	
		my $portID = $1;
		my $portGbic = $5;
		
		push(@portIDs,$portID);
		
		$PORT_DATA{"$portID#PORTGBIC#PORT"} = $portGbic;
		
		#print " $portID..........$portGbic\n";
	
	}


	### MLT config in CLI Mode 
	
	#mlt 11 create
		if ($_ =~ /mlt\s+(\d+)\s+create/){
	
			$mltID = $1;
		
			push(@mltIDs,$mltID);
			
		}
	#mlt 11 add ports 1/1,10/1
		
		if ($_ =~ /mlt\s+(\d+)\s+add\s+ports\s+(.*)/){
		
			$mltPorts = $2;
		
		$MLT_DATA{"MLT#$mltID#MLTPORT"}=$mltPorts;
	}	
	#mlt 11 name "SMLT-GovtCtr(10G)" 
				
		if ($_ =~ /mlt\s+(\d+)\s+name\s+(.*)/){
	
			$mltName = $2;
		
		$MLT_DATA{"MLT#$mltID#MLTNAME"}=$mltName;
	   }
	#mlt 11 smlt create smlt-id 4
	
		if ($_ =~ /mlt\s+(\d+)\s+smlt\s+create\s+smlt-id\s+(\d+)/){
		
			$mltType =  "smlt";
		
		$MLT_DATA{"MLT#$mltID#MLTTYPE"}=$mltType;
		
		}
	#mlt 13 ist enable
	
		if ($_ =~ /mlt\s+(\d+)\s+ist\s+enable/){
		
			$mltType =  "IST";
		
		$MLT_DATA{"MLT#$mltID#MLTTYPE"}=$mltType;
		
		}
	
	### MLT config in ACLI Mode 
	
	#mlt 1 enable name "DMLT ESSB to OSB"

		if ($_ =~ /mlt\s+(\d+)\s+enable\s+name\s+(.*)/){
		
			$mltID =  $1;
			$mltName = $2;
			
			push(@mltIDs,$mltID);
		
		$MLT_DATA{"MLT#$mltID#MLTNAME"}=$mltName;
		#print "MLT IDs : $mltID\n";
		#print "MLT IDs : $mltName\n";		
		}
	#mlt 1 member 7/16,7/32,8/16,8/32
	
		if ($_ =~ /mlt\s+(\d+)\s+member\s+(.*)/){
		
			$mltPorts = $2;
			
			$MLT_DATA{"MLT#$mltID#MLTPORT"}=$mltPorts;
		#print "MLT IDs : $mltPorts\n";
		}
	# int mlt 1
	# smlt 1
	# ist enable
	
		if ($_ =~ /interface\s+mlt\s+(\d+)/){
		
				$mltID =  $1;
		}
	
		if ($_ =~ /smlt\s+(\d+)/){
		
			$mltType =  "SMLT";		
			$MLT_DATA{"MLT#$mltID#MLTTYPE"}=$mltType;			
		
		}
		if ($_ =~ /ist\s+enable/){
		
			$mltType =  "IST";		
			$MLT_DATA{"MLT#$mltID#MLTTYPE"}=$mltType;

			 #print "MLT Types : $mltType\n";
		
		}
		 
	
		

	
	if (/                                   Port Name/ .. /                                   Vlan Port/) {
			
				#2/7   SW1K1                Gbic1310(Lx)  up       full     1000     Tagged
				
				if ($_ =~ /((\d+)\/(\d+))\s+(.*?)\s+(.*?)\s+(up|down)\s+(full|half)\s+(\d+)\s+(\w+)/){
				
					my $portID = $1;
					my $portName = $4;
					my $portDesc = $5;
					my $portState = $6;
					my $portDuplex = $7;
					my $portSpeed = $8;
					my $portTagging = $9;
					
					
					push(@portIDs,$portID);
					
						
						$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
						$PORT_DATA{"$portID#PORTDUPLEX#PORT"}=$portDuplex;
						$PORT_DATA{"$portID#PORTSPEED#PORT"}=$portSpeed;
						$PORT_DATA{"$portID#PORTTAG#PORT"}=$portTagging;
					
						#$PORT_DATA{"$portID#PORTNAME#PORT"}=$portName;
						
						#print " $portID..........$portState\n";
	
					
					#print " $portID ...$portState....$portDuplex...$portSpeed....$portTagging\n";
					
					
				}
				
				
	}
	
	### Getting the Port Name for the config file for ERS Switch  
	
	#ethernet 2/5 name "SMLT SW121 2/5 to ESSB 7/5"
	
	
	if ($_ =~ /ethernet\s+((\d+)\/(\d+))\s+name\s+(.*)/){
				
					my $portID = $1;
					my $portName = $4;

					push(@portIDs,$portID);
					
					$PORT_DATA{"$portID#PORTNAME#PORT"}=$portName;
					#print " $portID..........$portName\n";
					
	}
	
	
	### Port Name from Config for VOSS 
	if (/# PORT CONFIGURATION - PHASE II/ .. /# IP CONFIGURATION/) {	
	
		#interface GigabitEthernet 2/41
	
		
			if ($_ =~ /interface\s+GigabitEthernet\s+((\d+)\/(\d+))\s+/){
		
						$portID = $1;
						$portName = "";
						$PORT_DATA{"$portID#PORTNAME#PORT"}=$portName;
						push(@portIDs,$portID);
			}
			
			#name "MICe-Layer2"
			if ($_= /name\s+(.*)/){
				
					$portName = $1;
					
					#push(@portIDs,$portID);
					$PORT_DATA{"$portID#PORTNAME#PORT"}=$portName;
					
					
					
			
			}

	}
	### Works for ERS 8K / VOSS etc 
	if (/                                   Port Vlans/ .. /                              Port VRF Association/) {
	
		#2/5   enable  false   false     1       1 60 70 80 110   disable 
	
		if ($_ =~ /((\d+)\/(\d+))\s+(enable|disable)\s+(false|true)\s+(false|true)\s+(\d+)\s+(.*?)\s+(disable|enable)/){
		
				my $portID = $1;
				my $portVlans = $8;
	
				push(@portIDs,$portID);
					
				$PORT_DATA{"$portID#PORTVLANS#PORT"}=$portVlans;
		
		}

	
	}
	# 1/1   disable disable   7       disable 7  (For ERS 8300)
		if (/                                   Port Vlans/ .. /                            Port Unknown-Mac-Discard/) {
	
	
		if ($_ =~ /((\d+)\/(\d+))\s+(enable|disable)\s+(enable|disable)\s+(\d+)\s+(disable|enable)\s+(.*)\s+/){
		
				my $portID = $1;
				my $portVlans = $8;
				
				#print "$portID............$portVlans\n";
	
				push(@portIDs,$portID);
					
				$PORT_DATA{"$portID#PORTVLANS#PORT"}=$portVlans;
		
		}

	
	}
	

	
#############################################################################################
############################### EOS Parsing
#############################################################################################
	
	### set vlan create 5,10,31,100,1040-1047,2040,2140,3040-3041,3140-3146,3240,3340,3410,4004,4040,4056
	### Extracting the VLAN ID from the config and just putting a "blank for the Name Section"
	
	
	if ($_=~/set\s+vlan\s+create\s+(.*)/){
	
			
		$vlanNumberList = $1;
		(@vlanNumberList) = convert_to_full_vlan_list($vlanNumberList);
		
		foreach $vlanID(@vlanNumberList)
		{
			print "$vlanID\n";
			$vlanID = trim($vlanID);
			push(@vlanIDs,$vlanID);
			
			$VLAN_DATA{"$vlanID#VLANNAME#VLAN"}="";		
		}	
	}
	
	#set vlan name 5 vaxip
	if ($_=~/set\s+vlan\s+name\s+(\d+)\s+(.*)/){
		
		$vlanID = $1;
		$vlanName = $2;
		$VLAN_DATA{"$vlanID#VLANNAME#VLAN"}=$vlanName;

	}
	
	
	
	#VLAN Interface
	##interface vlan.0.1040
	if($_ =~ /^\s*interface\s+vlan\.(\d+)\.(\d+)/){
    	$inInterfaceConfig = 1;
		$vlanID = $2;
		$vlanID = trim($vlanID);
				
    }
	if($_ =~ /exit/){
		if ($inInterfaceConfig){
			#printVlanMember();
			$inInterfaceConfig = 0;
    clearVars();
    #print $line." :::".$1."\n";#."\n";
		}
	}

	## IP Add and Mas 
	## ip address 192.204.144.97 255.255.255.224 primary
	if($_ =~ /^\s*ip\s+address\s+(\d+.\d+.\d+.\d+)\s+(\d+.\d+.\d+.\d+)\s+primary/){
   
		$vlanIP = $1;
		$vlanMask = $2;
		$VLAN_DATA{"$vlanID#VLANIP#VLAN"}=$vlanIP;
		$VLAN_DATA{"$vlanID#VLANMASK#VLAN"}=$vlanMask;
   }
   ###ip address 172.17.41.1 255.255.255.0 secondary
	if($_ =~ /^\s*ip\s+address\s+(\d+.\d+.\d+.\d+)\s+(\d+.\d+.\d+.\d+)\s+secondary/){
   
		$vlanIP = $1;
		$vlanMask = $2;
		$VLAN_DATA{"$vlanID#VLANIP#VLAN"} .= " # $vlanIP";
		$VLAN_DATA{"$vlanID#VLANMASK#VLAN"} .= " # $vlanMask";
   }
   ## ip helper-address 198.17.40.11
	if($_ =~ /ip\s+helper\-address\s+(\d+.\d+.\d+.\d+)/){
		
		$serverIP = $1;
		$VLAN_DATA{"$vlanID#VLANIP#DHCPSERVER"} .= ",$serverIP";
		
	}
	#  vrrp create 5 v3-IPv4
	#  vrrp address 5 198.17.40.12 
	#  vrrp accept-mode 5
	#  vrrp enable 5
	if($_ =~ /vrrp\s+address\s+(\d+)\s+(\d+.\d+.\d+.\d+)/){
		
		$vlanVrrpId = $1;
		$vlanVrrpIp = $2;
		#print "$1...$2\n";
		$VLAN_DATA{"$vlanID#VLANVRRPID#VRRPID"} = $vlanVrrpId;
		$VLAN_DATA{"$vlanID#VLANVRRPIP#VRRPIP"} = $vlanVrrpIp;
		
	}	



##ge.1.1       NewDorm2         down     up        1.0G full    1000-lx      lc

	
	if ($_ =~ /(ge\.(\d+)\.(\d+))\s+(.*?)\s+(up|down|dormant)\s+(up|down)\s+(.*?)\s+(full|half)\s+(.*)/){
	
			my $portID = $1;
			my $portName = $4;
			my $portState = $5;
			my $portDuplex = $8;
			my $portSpeed = $7;
			my $portGbic = $9;
			
			
			push(@portIDs,$portID);
			
				$PORT_DATA{"$portID#PORTNAME#PORT"}=$portName;
				$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
				$PORT_DATA{"$portID#PORTDUPLEX#PORT"}=$portDuplex;
				$PORT_DATA{"$portID#PORTSPEED#PORT"}=$portSpeed;
				$PORT_DATA{"$portID#PORTGBIC#PORT"} = $portGbic;
	
	}
	
##tg.1.3       10G PtoP Pfahler up       up       10.0G full    10g-lr       lc

	
	if ($_ =~ /(tg\.(\d+)\.(\d+))\s+(.*?)\s+(up|down|dormant)\s+(up|down)\s+(.*?)\s+(full|half)\s+(.*)/){
	
		#print "$1...$5...$7..$8..$9\n";
			my $portID = $1;
			my $portName = $4;
			my $portState = $5;
			my $portDuplex = $8;
			my $portSpeed = $7;
			my $portGbic = $9;
			
			
			push(@portIDs,$portID);
			
				$PORT_DATA{"$portID#PORTNAME#PORT"}=$portName;
				$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
				$PORT_DATA{"$portID#PORTDUPLEX#PORT"}=$portDuplex;
				$PORT_DATA{"$portID#PORTSPEED#PORT"}=$portSpeed;
				$PORT_DATA{"$portID#PORTGBIC#PORT"} = $portGbic;
	
	}
	
	#set vlan egress 5 lag.0.3-4;ge.1.1-2,8,10,15;tg.1.4 tagged
		if ($_=~/set\s+vlan\s+egress\s+(\d+)\s+(.*?)\s+(tagged|untagged)/){
	
		my $vlanID = $1;
		my  $portList = $2;
		
		#print ".......$2..\n";
		
		
		($returnPortList) = convert_to_full_port_list($portList);
				
				foreach $portID  (split(/\,/,$returnPortList)){
					$portID = trim ($portID);
					
						print "****$portID\n";
					
						if (($PORT_DATA{"$portID#PORTVLANS#PORT"}) eq "") {
						
							$PORT_DATA{"$portID#PORTVLANS#PORT"}=$vlanID;
							#print "****$portID************$vlanID\n";
													
						}else {
						
							$PORT_DATA{"$portID#PORTVLANS#PORT"}.= " ,$vlanID";
			
						}

				}
				
		
	
	
	}
	
	
	
	
#############################################################################################
############################### EXOS Parsing
#############################################################################################

	

	
	### configure vlan Default ipaddress 10.1.2.23 255.255.0.0
	### configure vlan apparatus ipaddress 10.14.0.21 255.255.255.252
	
	
	if ($_=~/configure\s+vlan\s+(.*)\s+ipaddress\s+(\d+.\d+.\d+.\d+)\s+(\d+.\d+.\d+.\d+)/){
	
		my $vlanName = $1;
		my $vlanIP = $2;
		my $vlanMask = $3;
		
				
		
		@uniqVlanNames = uniq @vlanNames;                   #######Using this step to make the array unique
		@uniqVlanNames = grep /\S/, @uniqVlanNames;         ### cleanup of array for any empty name etc 
			foreach $uniqVlanName (@uniqVlanNames) {
			
				if ( $vlanName eq $uniqVlanName ){ 	

					my $vlanID  = $VLAN__DATA_NAME{"$vlanName#VLANID"};   ### Looking up the VLAN ID for the Name
				
						$VLAN_DATA{"$vlanID#VLANIP#VLAN"}=$vlanIP;
						$VLAN_DATA{"$vlanID#VLANMASK#VLAN"}=$vlanMask;
						$VLAN_DATA{"VLAN_BY_IP#$vlanIP"}=$vlanID;					
				}				
			}
	}
	
	### VRRP## create vrrp vlan OVS_vRP-MGMT vrid 30
	
	if ($_=~/create\s+vrrp\s+vlan\s+(.*)\s+vrid\s+(\d+)/){
	
	
			my $vlanName = $1;
			my $vlanVrrpId = $2;
		
			push(@vlanIDs,$vlanID);
			
		@uniqVlanNames = uniq @vlanNames;                   #######Using this step to make the array unique
		@uniqVlanNames = grep /\S/, @uniqVlanNames;         ### cleanup of array for any empty name etc 
			foreach $uniqVlanName (@uniqVlanNames) {
			
				if ( $vlanName eq $uniqVlanName ){ 		      
					my $vlanID  = $VLAN__DATA_NAME{"$vlanName#VLANID"};   ### Looking up the VLAN ID for the Name
				
						$VLAN_DATA{"$vlanID#VLANVRRPID#VRRPID"} = $vlanVrrpId;
		
				
				}				
			}
	}
	
	### VRRP VRRPIP #configure vrrp vlan OVS_vRP-MGMT vrid 30 add 20.20.73.1
		
	if ($_=~/configure\s+vrrp\s+vlan\s+(.*)\s+vrid\s+(\d+)\s+add\s+(\d+.\d+.\d+.\d+)/){
	
	
			my $vlanName = $1;
			my $vlanVrrpId = $2;
			my $vlanVrrpIp = $3;
		
			push(@vlanIDs,$vlanID);
			
		@uniqVlanNames = uniq @vlanNames;                   #######Using this step to make the array unique
		@uniqVlanNames = grep /\S/, @uniqVlanNames;         ### cleanup of array for any empty name etc 
			foreach $uniqVlanName (@uniqVlanNames) {
			
				if ( $vlanName eq $uniqVlanName ){ 		      
					my $vlanID  = $VLAN__DATA_NAME{"$vlanName#VLANID"};   ### Looking up the VLAN ID for the Name
				
						$VLAN_DATA{"$vlanID#VLANVRRPIP#VRRPIP"} = $vlanVrrpIp;;
		
				
				}				
			}
	}
	
	## Checking if DHCP Relay is enable Globally 	# configure bootprelay add 10.1.2.14 vr VR-Default
				

	if ($_=~/configure\s+bootprelay\s+add\s+(\d+.\d+.\d+.\d+)\s+vr\s+VR-Default/){
	
			$bootIP = $1; 
			push(@serverIPS,$bootIP);    ## Pushing Server IP's in an array , so that I will make it unique later and copy the content of array into String
			@serverIPS = uniq @serverIPS;   ## Making the array unique to remove duplicate items 
			$serverIP = join ",", @serverIPS;   ## Converting the array to string , as I push the values in Hash for that I need a string 
		}
	
		
	
	## Check which VLAN has DHCP enable and get the value from the global  #enable bootprelay ipv4 vlan VLAN_0010
	
	
	if ($_=~/enable\s+bootprelay\s+ipv4\s+vlan\s+(.+?)\s*$/){
		
		my $vlanName = $1;

		@uniqVlanNames = uniq @vlanNames;                   #######Using this step to make the array unique
		@uniqVlanNames = grep /\S/, @uniqVlanNames;         ### cleanup of array for any empty name etc 
			foreach $uniqVlanName (@uniqVlanNames) {
			
				if ( $vlanName eq $uniqVlanName ){ 		      
					my $vlanID  = $VLAN__DATA_NAME{"$vlanName#VLANID"};   ### Looking up the VLAN ID for the Name
				
						$VLAN_DATA{"$vlanID#VLANIP#DHCPSERVER"} = $serverIP;
	
				}				
			}
	
	
	}
	
	## Check if DHCP relay is configured  on specific vlan   ##configure bootprelay vlan Line_Dept_102 add 2.2.2.2
	
	
		if ($_=~/configure\s+bootprelay\s+vlan\s+(.*)\s+add\s+(\d+.\d+.\d+.\d+)/){
		
			my $vlanName = $1;
			my $bootpIP = $2;


			@uniqVlanNames = uniq @vlanNames;                   #######Using this step to make the array unique
			@uniqVlanNames = grep /\S/, @uniqVlanNames;         ### cleanup of array for any empty name etc 
				foreach $uniqVlanName (@uniqVlanNames) {
				
					if ( $vlanName eq $uniqVlanName ){ 		      
						my $vlanID  = $VLAN__DATA_NAME{"$vlanName#VLANID"};   ### Looking up the VLAN ID for the Name
					
							$VLAN_DATA{"$vlanID#VLANIP#DHCPSERVER"} .= " ,$bootpIP";
			
					}				
				}
			
	
	}
	
	
	# MLT ## enable sharing 1:7 grouping 1:7,2:7 algorithm address-based L2 lacp
	  ## enable sharing 49 grouping 49 algorithm address-based L3_L4 lacp
	  ## enable sharing 10 grouping 10-11 algorithm roundRobin-based
	  

		if ($_=~/enable\s+sharing\s+(.*)\s+grouping\s+(.*)\s+algorithm\s+(.*)/){
		
				$mltID =  $1;
				$mltPorts = $2;
			
				push(@mltIDs,$mltID);
		
					$MLT_DATA{"MLT#$mltID#MLTPORT"}=$mltPorts;
			

		}
		
		
		
		
	#1:1      VR-Default  E      A    ON  AUTO  1000 AUTO FULL SY/ASY       UTP
	if ($_=~/(.*?)\s+VR-Default\s+(\w+)\s+(\w+)\s+(ON|OFF|NONE)\s+(AUTO|\d+)\s+(\d+)\s+(AUTO|FULL)\s+(FULL|HALF|AUTO)\s+(SYM|NONE|ASY|SY\/ASY)\s+(.*)/){
	
		#print "In..$1 ...$2....$3....$6....$10\n";
		
					
					my $portID = $1;
					my $portDuplex = $8;
					my $portSpeed = $6;
					my $portGbic = $10;
					
				
					$portID = trim($portID);
					
					if ( $3 eq "A") {
					    my 	$portState = "up";
						$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
					}
					
					if (( $2 eq "D") | ($3 eq "R") |  ($3 eq "NP")){
					    my 	$portState = "down";
						$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
					}
						

					if ( $portGbic =~/(\d+):(\d+)\s+(.*?)/){     #### Removing Load Master from the o/p  e.g :21  LX  UTP
						
						$portGbic =~ s/(\d+):(\d+)//;     ### Using perl regex substitution 
					}
							

					
					push(@portIDs,$portID);
					
						
						#$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
						$PORT_DATA{"$portID#PORTDUPLEX#PORT"}=$portDuplex;
						$PORT_DATA{"$portID#PORTSPEED#PORT"}=$portSpeed;
						$PORT_DATA{"$portID#PORTGBIC#PORT"} = $portGbic;					
				
					
	}
	
	
	
	#2:23     VR-Default  E     NP    ON  AUTO       AUTO                    NONE
	#1:4      VR-Default D       R    ON  AUTO       AUTO                    UTP
	#1        VR-Default  E      R   OFF 10000       FULL                    Q+SR4
	#STAFF02  VR-Default  E      R   OFF 10000       FULL                    BASET
	
		if ($_=~/(.*?)\s+VR-Default\s+(\w+)\s+(\w+)\s+(ON|OFF|NONE)\s+(AUTO|\d+)\s+(AUTO|FULL)\s+(.+?)\s*$/){
	
		#print "In..$1...$2...$3....$7 \n";
		
					my $portID = $1;
					my $portDuplex = "0";
					my $portSpeed = "0";
					my $portGbic = $7;
					$portID = trim($portID);
					
					if ( $3 eq "A") {
					    my 	$portState = "up";
						$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
					}
			
					if (( $2 eq "D") |  ($3 eq "R") |  ($3 eq "NP"))  {
					    my 	$portState = "down";
						$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
					}
						
					
					push(@portIDs,$portID);
					
						
						#$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
						$PORT_DATA{"$portID#PORTDUPLEX#PORT"}=$portDuplex;
						$PORT_DATA{"$portID#PORTSPEED#PORT"}=$portSpeed; 
						$PORT_DATA{"$portID#PORTGBIC#PORT"} = $portGbic;
	
	
	    }
	
		#Firehall>VR-Default  E      A   OFF 10000 10000 FULL FULL   NONE        SF+_LR 
		#CLASS01_>VR-Default  E      R   OFF 10000       FULL                    NONE    
		#VAULT01_>VR-Default  E      A   OFF 10000 10000 FULL FULL   NONE        SF+_CX3m
		#PCRF-GX->VR-Default  E      A   OFF 10000 10000 FULL FULL   NONE        Q+SR4
		if ($_=~/(.*?)\>VR-Default\s+(\w+)\s+(\w+)\s+(ON|OFF|NONE)\s+(AUTO|\d+)\s+(\d+)\s+(AUTO|FULL)\s+(FULL|HALF|AUTO)\s+(SYM|NONE|ASY|SY\/ASY)\s+(.+?)\s*$/){
	
		#print "In..$1...$2...$3....$10 \n";
		
					
		
					my $portID = $1;
					my $portDuplex = $8;
					my $portSpeed = $6;
					my $portGbic = $10;
					$portID = trim($portID);
					
			
					if ( $3 eq "A") {
					    my 	$portState = "up";
						$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
					}
			
					if (( $2 eq "D") |  ($3 eq "R")  |  ($3 eq "NP")) {
					    my 	$portState = "down";
						$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
					}
						
					
					push(@portIDs,$portID);
					#print "...IN...$portID\n";
						
					#	$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
						$PORT_DATA{"$portID#PORTDUPLEX#PORT"}=$portDuplex;
						$PORT_DATA{"$portID#PORTSPEED#PORT"}=$portSpeed; 
						$PORT_DATA{"$portID#PORTGBIC#PORT"} = $portGbic; 
	
	
	}
	
	#CLASS01_>VR-Default  E      R   OFF 10000       FULL                    NONE  
		if ($_=~/(.*?)\>VR-Default\s+(\w+)\s+(\w+)\s+(ON|OFF|NONE)\s+(AUTO|\d+)\s+(AUTO|FULL)\s+(.+?)\s*$/){
		
					my $portID = $1;
					$portID = trim($portID);
					

	
					if ( $3 eq "A") {
					    my	$portState = "up";
							$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
					}
			
					if (( $2 eq "D") |  ($3 eq "R")  |  ($3 eq "NP")) {
					    my 	$portState = "down";
							$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
					}
						
					
					push(@portIDs,$portID);
					
						
						#$PORT_DATA{"$portID#PORTSTATE#PORT"}=$portState;
						$PORT_DATA{"$portID#PORTDUPLEX#PORT"}=$portDuplex;
						$PORT_DATA{"$portID#PORTSPEED#PORT"}=$portSpeed; 
						$PORT_DATA{"$portID#PORTGBIC#PORT"} = $portGbic; 
	
	
	}
	
	
	######## POrt to VLAN Mapping 
	#configure vlan LTE-SGI-L3SW-17-01 add ports 11 untagged
	#configure vlan OVS_vSFO-3 add ports 77-79,81-83,85-86,103 tagged
	if ($_=~/configure\s+vlan\s+(.*?)\s+add\s+ports\s+(.*?)\s+(tagged|untagged)/){
	
	
		my $vlanName = $1;
		my  $portList = $2;
		
		
			foreach $uniqVlanName (@vlanNames) {
				#print "\n*******$uniqVlanName********\n";
				if ( $vlanName eq $uniqVlanName ){ 
					
					$vlanID  = $VLAN__DATA_NAME{"$vlanName#VLANID"};   ### Looking up the VLAN ID for the Name
					#print "\n$vlanID..$vlanName...$portList...OUT..";
				}
		}
		
		(@portList) = convert_to_full_port_list($portList);
				
				foreach $portID (@portList) {
				
					$portID = trim ($portID);
					
					print "$portID";
				
					if (($PORT_DATA{"$portID#PORTVLANS#PORT"}) eq "") {
					
						$PORT_DATA{"$portID#PORTVLANS#PORT"}=$vlanID;
						
																		
					}else {
					
						$PORT_DATA{"$portID#PORTVLANS#PORT"}.= " ,$vlanID";
						print "                              																$vlanID\n";
		
					}				

				}
				
	}
	
	
}


############## This is the logic where I'm replacing the port Name with the Port Number , this is a important logic for EXOS ,as they trunketed the port name in display O/P


		foreach $portID (@portIDs){
		
			my $PORT_STATE  =$PORT_DATA{"$portID#PORTSTATE#PORT"};
			my $PORT_DUPLEX  =$PORT_DATA{"$portID#PORTDUPLEX#PORT"};
			my $PORT_SPEED  =$PORT_DATA{"$portID#PORTSPEED#PORT"};
			my $PORT_TAG = $PORT_DATA{"$portID#PORTTAG#PORT"};
			my $PORT_NAME  =$PORT_DATA{"$portID#PORTNAME#PORT"};
			my $PORT_GBIC  =$PORT_DATA{"$portID#PORTGBIC#PORT"};
			my $PORT_VLANID  = $PORT_DATA{"$portID#PORTVLANS#PORT"};
	    
			if ( $portID !~ /^((\d+):(\d+))|^(\d+)/ ) { 				### Negative Regex , which means if port has Name 
			
												
				foreach $portDesc (@portNames) {
					chomp ();
					 		      
						my $pID  = $PORT__DATA_NAME{"$portDesc#PORTID"};   ### Looking up the PORT ID for the Name
						$portID = $pID;
					
						push(@portIDs,$portID) ; ## Here I'm pushing the new port ID as another Hash Key and will set all the same value for the partial name
							
							$PORT_DATA{"$portID#PORTDUPLEX#PORT"}=$PORT_DUPLEX;
							$PORT_DATA{"$portID#PORTSPEED#PORT"}=$PORT_SPEED; 
							$PORT_DATA{"$portID#PORTGBIC#PORT"} = $PORT_GBIC;
							$PORT_DATA{"$portID#PORTSTATE#PORT"}=$PORT_STATE;
							$PORT_DATA{"$portID#PORTNAME#PORT"}=$portDesc;
							$PORT_DATA{"$portID#PORTVLANS#PORT"} = $PORT_VLANID;
							#	print "$portID\n";
							#	print $PORT_DATA{"$portID#PORTVLANS#PORT"};
							#if ($portID =~/\,/){   ### Matching if there is a comma in the Port ID which means same port name for multiple ports 
							#
							#	my ($portID, $portIDNext) = split /,/, $portID;   ##Spiliting the duplicate port ID's and getting the first ID 
							#
							#	#print " $portID ....$PORT_SPEED \n";
							#
							#}
						
				
							shift @portNames;     ### Here getting the first entry out of the arrary , so that the next time the foreach loop runs , first entry is removed
							last ;          #### Last means to break out of the inner for loop  
			    
					    			
				}
                
			}
			
					
			
		}

		

		
### This part of code assign the VLAN Membership to all ports of an MLT for EXOS
		
		
	@mltIDs = uniq @mltIDs;
	
	foreach $mltID (@mltIDs) {
	
		my $MLT_PORTS  = $MLT_DATA{"MLT#$mltID#MLTPORT"};   ### Assigning into an array 
		my @MLT_PORTS = split /,/, $MLT_PORTS;
		my @MLT_PORTS = split /-/, $MLT_PORTS;        ###In some case the MLT ports are mapped like 10-11
		my $portID = $mltID;
		my $PORT_VLANID = $PORT_DATA{"$portID#PORTVLANS#PORT"};
		
			foreach $portID (@MLT_PORTS){
			
				$PORT_DATA{"$portID#PORTVLANS#PORT"} = $PORT_VLANID;
			
			}        		
	
	}
	
#### This part of code goes over the EXOS port list and 
#
#	@uniqportIDs = uniq @portIDs;
#	
#	@uniqportIDs = grep /\S/, @uniqportIDs;
#	
#		foreach $portID (@uniqportIDs) {
#		
#			print "***$portID*****\n";
#		
#		
#		}
#	
	

################################	PRINTING SECTION ##############################
	
	$worksheet->write( $rowcounter, 0, , "Hostname", $format_purple );
    $worksheet->write( $rowcounter, 1, , "VLAN ID", $format_purple );
	$worksheet->write( $rowcounter, 2, , "VLAN Name", $format_purple );
	$worksheet->write( $rowcounter, 3, , "I-SID", $format_purple );
	$worksheet->write( $rowcounter, 4, , "VRF Name", $format_purple );
	$worksheet->write( $rowcounter, 5, , "VLAN IP", $format_purple );
	$worksheet->write( $rowcounter, 6, , "Subnet Mask", $format_purple );
	$worksheet->write( $rowcounter, 7, , "VRID", $format_purple );
	$worksheet->write( $rowcounter, 8, , "VRRP IP", $format_purple );
	$worksheet->write( $rowcounter, 9, , "RSMLT Edge (Type Y)", $format_purple );
	$worksheet->write( $rowcounter, 10, , "DHCP Relay", $format_purple );
	$worksheet->write( $rowcounter, 12, , "MLT ID", $format_purple );
	$worksheet->write( $rowcounter, 13, , "MLT Name", $format_purple );
	$worksheet->write( $rowcounter, 14, , "Port No", $format_purple );
	$worksheet->write( $rowcounter, 15, , "SMLT/ NNI/      FA-SMLT/ MLT", $format_purple );
	$worksheet->write( $rowcounter, 16, , "FA Mgmt (vlan,I-SID)", $format_purple );
	
	$worksheet->write( $rowcounter, 18, , "Port No", $format_purple );
	$worksheet->write( $rowcounter, 19, , "New Port", $format_yellow );
	$worksheet->write( $rowcounter, 20, , "Port Name", $format_purple );
	$worksheet->write( $rowcounter, 21, , "VLAN IDs", $format_purple );
	$worksheet->write( $rowcounter, 22, , "New VLAN IDs",$format_yellow );
	$worksheet->write( $rowcounter, 23, , "SMLT/   NNI/        FA-SMLT/ MLT/ UNI",$format_purple );
	$worksheet->write( $rowcounter, 24, , "Speed", $format_purple );
	$worksheet->write( $rowcounter, 25, , "Status", $format_purple );
	$worksheet->write( $rowcounter, 26, , "Port Type", $format_purple );
	
	$rowcounter++;
	
	@uniqportIDs = uniq @portIDs;
	
	@uniqportIDs = grep /\S/, @uniqportIDs;
	
		foreach $portID (@uniqportIDs) {
		
			my $PORT_STATE  =$PORT_DATA{"$portID#PORTSTATE#PORT"};
			my $PORT_DUPLEX  =$PORT_DATA{"$portID#PORTDUPLEX#PORT"};
			my $PORT_SPEED  =$PORT_DATA{"$portID#PORTSPEED#PORT"};
			my $PORT_TAG = $PORT_DATA{"$portID#PORTTAG#PORT"};
			my $PORT_NAME  =$PORT_DATA{"$portID#PORTNAME#PORT"};
			my $PORT_GBIC  =$PORT_DATA{"$portID#PORTGBIC#PORT"};
			my $PORT_VLANID  = $PORT_DATA{"$portID#PORTVLANS#PORT"};
			
			#print "$PORT_VLANID\n";
			$portID = trim ($portID);
			$PORT_NAME =~ s/^\s+|\s+$//g;  ### Removing leading and trailing space 
			$PORT_NAME =~ s/^"|"$//g;   ### Removing First and Last Inverted comma
			$PORT_VLANID = trim ($PORT_VLANID);
			
				$worksheet->write($rowcounter, 18,  $portID ,$format_def);
				$worksheet->write($rowcounter, 19, "" ,$format_def);
				$worksheet->write($rowcounter, 20,  $PORT_NAME ,$format_def);
				$worksheet->write($rowcounter, 21,  $PORT_VLANID ,$format_def);
				$worksheet->write($rowcounter, 22, "" ,$format_def);
				$worksheet->write($rowcounter, 23, "" ,$format_def);
				$worksheet->write($rowcounter, 24,  $PORT_SPEED ,$format_def);
				$worksheet->write($rowcounter, 25, $PORT_STATE ,$format_def);
				$worksheet->write($rowcounter, 26,  $PORT_GBIC ,$format_def);
				
				#print " $rowcounter.********$portID................$PORT_NAME\n";
								
				$rowcounter++;		
		}
	
		$rowcounter = "1";
		
	
	
	
	
	@mltIDs = uniq @mltIDs;
	
	foreach $mltID (@mltIDs) {
	
		my $MLT_NAME  = $MLT_DATA{"MLT#$mltID#MLTNAME"};
		my $MLT_PORT  = $MLT_DATA{"MLT#$mltID#MLTPORT"};
		my $MLT_TYPE  = $MLT_DATA{"MLT#$mltID#MLTTYPE"};
						
		
		#$MLT_NAME =~ s/^,|,$//g;   ### Removing First and Last comma 
	
				$worksheet->write($rowcounter, 12,  $mltID ,$format_def);
				$worksheet->write($rowcounter, 13, $MLT_NAME ,$format_def);
				$worksheet->write($rowcounter, 14, $MLT_PORT ,$format_def);
				$worksheet->write($rowcounter, 15, $MLT_TYPE ,$format_def);
				$worksheet->write($rowcounter, 16, "" ,$format_def);
	
				$rowcounter++;	
	
	}
	
	
	$rowcounter = "1";
	
	
		
	
	@uniqvlanIDs = uniq @vlanIDs;
	
	foreach $vlanID (@uniqvlanIDs) {
	
		my $VLAN_NAME  = $VLAN_DATA{"$vlanID#VLANNAME#VLAN"};
		my $VLAN_IP  = $VLAN_DATA{"$vlanID#VLANIP#VLAN"};
		my $VLAN_MASK  = $VLAN_DATA{"$vlanID#VLANMASK#VLAN"};
		my $VLAN_DHCPSRV = $VLAN_DATA{"$vlanID#VLANIP#DHCPSERVER"};
		my $VLAN_VRRPID =$VLAN_DATA{"$vlanID#VLANVRRPID#VRRPID"};
		my $VLAN_VRRPIP = $VLAN_DATA{"$vlanID#VLANVRRPIP#VRRPIP"};
		
		$VLAN_DHCPSRV =~ s/^,|,$//g;   ### Removing First and Last comma 
		
		
		
				$worksheet->write($rowcounter, 1,  $vlanID ,$format_def);
				$worksheet->write($rowcounter, 2, $VLAN_NAME ,$format_def);
				$worksheet->write($rowcounter, 5, $VLAN_IP ,$format_def);
				$worksheet->write($rowcounter, 6, $VLAN_MASK ,$format_def);
				$worksheet->write($rowcounter, 7, $VLAN_VRRPID ,$format_def);
				$worksheet->write($rowcounter, 8, $VLAN_VRRPIP ,$format_def);
				$worksheet->write($rowcounter, 10, $VLAN_DHCPSRV ,$format_def);
		
		
		#print "VLAN ID  : $vlanID......VLAN Name  : $VLAN_NAME ....VLAN IP : $VLAN_IP ...$VLAN_DHCPSRV\n";
	
	
				$rowcounter++;	
	
	}
	
sub convert_to_full_port_list
{
   my $port_string; my $port;
   ($port_string,@dumb) = @_;
   my $port_group; my @ret_ports;
   my %PORT_TYPE;

   my $new_port_string = "";
   foreach my $card_group (split(/\;/,$port_string))
   {
      if ($card_group =~ /^((ge)|(tg))\.(\d+)\./)
      {
         my $type = $1;
         my $this_card = $4;
         $card_group =~ s/^ge\.\d+\.//;
         $card_group =~ s/^tg\.\d+\.//;
         foreach my $port_group (split(/\,/,$card_group))
         {
            $new_port_string .= $this_card."\/".$port_group.",";
            $PORT_TYPE{"TYPE#GROUP#".$this_card."\/".$port_group} = $type;
            $PORT_TYPE{"STYLE#GROUP#".$this_card."\/".$port_group} = "EOS";
         }
      }
   }
   $new_port_string =~ s/\,$//;
   if ($new_port_string ne "")
   {
      $port_string = $new_port_string;
   }


   foreach $port_group (split(/\,/,$port_string))
   {
      my $port_type = "";
      my $port_style = "VOSS"; #Default style.
      my @group_ports;

      if ($PORT_TYPE{"TYPE#GROUP#$port_group"} ne "")
      {
         $port_type = $PORT_TYPE{"TYPE#GROUP#$port_group"};
      }
      if ($PORT_TYPE{"STYLE#GROUP#$port_group"} ne "")
      {
         $port_style = $PORT_TYPE{"STYLE#GROUP#$port_group"};
      }

      #Remove any padding spaces ie  0/ 0 to 0/0
      $port_group =~ s/\s+//g;

      if (/^\d+\:\d+/)
      {
         $port_style = "EXOS";
      }
      elsif (/^\d+\.\d+/)
      {
         $port_style= "EOS";
      }


      #Convert EXOS style to modular style
      #ie 4:2-5 will be 4:2-4:5
      #
      $port_group =~ s/^(\d+)\:(\d+)-(\d+)$/$1\:$2\-$1\:$3/;
	  
      #Convert EOS style to modular style
      #ie 4.2-5 will be 4.2-4.5
      #
      $port_group =~ s/^(\d+)\.(\d+)-(\d+)$/$1\.$2\-$1\.$3/;


      #Convert stack style to modular style
      #ie 4/2-5 will be 4/2-4/5
      #
      $port_group =~ s/^(\d+)\/(\d+)-(\d+)$/$1\/$2\-$1\/$3/;

      #Convert single stack style to modular style
      #ie  1-5,6,7,10-20 = 1/1-1/5,1/6,1/7,1/10-1/20

      #Ranges:
      $port_group =~ s/^(\d+)-(\d+)$/1\/$1\-1\/$2/;

      #Single ports:
      $port_group =~ s/^(\d+)$/1\/$1/;

      if ($port_group =~ /^(\d+)\/(\d+)-(\d+)\/(\d+)$/)
      {
         #Modular style 4/2-4/5 meaning 4/1,4/2,4/3,4/4,4/5
         my $card = $1; my $sport = $2; my $lport = $4;

         for(my $i=$sport;$i<$lport+1;$i++)
         {
            push(@ret_ports,"$card\/$i");
            push(@group_ports,"$card\/$i");
         }
      }
      elsif ($port_group =~ /^(\d+)\/(\d+)\/(\d+)-(\d+)\/(\d+)\/(\d+)$/)
      {
         my $card = $1; my $st_port = $2; my $st_sport = $3; my $lst_port = $4; my $lst_sport = $5;
         if ($st_port ne $lst_port)
         {
            #Modular sub-port style 1/41/1-1/42/4 meaning 1/41/1,1/41/2,1/41/3,1/41/4,1/42/1,1/42/2,1/42/3,1/42/4
            #However, don't want to assume only 4 sub ports, so need to attempt calc and then compare against ALL_PORTS_ARRAY
            #Will check for sub-ports 1 thru 128
            for(my $i=$st_port;$i<$lst_port+1;$i++)
            {
               for(my $j=$st_sport;$j<129;$j++)
               {
                  if ($PORT_DATA{"$card\/$i\/$j"})
                  {
                     push(@ret_ports,"$card\/$i\/$j");
                     push(@group_ports,"$card\/$i\/$j");
                  }
               }
            }
         }
         else
         {
            #Modular sub-port style 4/2/1-4/2/4 meaning 4/2/1,4/2/2,4/2/3,4/2/4
            for(my $i=$st_sport;$i<$lst_sport+1;$i++)
            {
               push(@ret_ports,"$card\/$st_port\/$i");
               push(@group_ports,"$card\/$st_port\/$i");
            }
         }
      }
      elsif ($port_group =~ /^(\d+)\/ALL$/i)
      {
         #Stackable style 1/ALL, 2/ALL, etc.
         my $card = $1;
         #Because the ALL_PORT_ARRAY is required for this
         #to work, it need to be populated first.
         foreach $port (split(/\,/,$PORT_DATA{"ALL_PORTS_ARRAY"}))
         {
            my $this_card; my $this_port; my $this_sub_port;
            ($this_card,$this_port,$this_sub_port) = split(/\//,$port);
            if ($card == $this_card)
            {
               push(@ret_ports,$port);
               push(@group_ports,$port);
            }
         }
      }
      elsif ($port_group =~ /^ALL$/i)
      {
         #Stackable style ALL
         #Because the ALL_PORT_ARRAY is required for this
         #to work, it need to be populated first.
         foreach $port (split(/\,/,$PORT_DATA{"ALL_PORTS_ARRAY"}))
         {
            push(@ret_ports,$port);
            push(@group_ports,$port);
         }
      }
      elsif ($port_group =~ /^NONE$/i)
      {
         #Stackable style NONE, used for the VLAN 1.
         #Do nothing.
      }
      else
      {
         push(@ret_ports,$port_group);
         push(@group_ports,$port_group);
      }

      foreach my $port (@group_ports)
      {
         $PORT_TYPE{"TYPE#PORT#$port"} = $port_type;
         $PORT_TYPE{"STYLE#PORT#$port"} = $port_style;
      }
   }

   my $ret_port_string = "";
   foreach my $port (@ret_ports)
   {
      my $type = $PORT_TYPE{"TYPE#PORT#$port"};
      my $style= $PORT_TYPE{"STYLE#PORT#$port"};
      $ret_port_string .= convert_port_type($port,$style,$type).",";
   }
   $ret_port_string =~ s/\,$//;
   return($ret_port_string);
   #return(join(",",@ret_ports));
}



#Only accepts single port as input.
#Can be of any style to start with.
#Will output in style specified.
sub convert_port_type
{
   my $port = shift;
   my $style = shift;
   my $type = shift;

   #Convert from EXOS to VOSS
   $port =~ s/^(\d+)\:(\d+)/$1\/$2/;
	  
   #Convert from EOS to VOSS
   $port =~ s/^((tg)|(ge))\.(\d+)\.(\d+)/$4\/$5/;

   #Convert from EOS(no type) to VOSS
   $port =~ s/^(\d+)\.(\d+)/$1\/$2/;

   #Convert from single digit(ie no stack) to VOSS
   $port =~ s/^(\d+)$/1\/$1/;

   if ($style eq "VOSS")
   {
      return($port);
   }
   elsif ($style eq "EXOS")
   {
      $port =~ s/\:/\//;
      return($port);
   }
   elsif($style eq "EOS")
   {
     my $card; my $c_port;
     ($card,$c_port) = split(/\//,$port);
     if ($type ne "" && $card ne "")
     {
        return("$type\.$card\.$c_port");
     }
     elsif ($card ne "")
     {
        return("$card\.$c_port");
     }
     else
     {
        return($port);
     }
   }

   return($port);
}

	
	
	
#sub convert_to_full_port_list
#{
#   my $port_string; my $port;
#   ($port_string,@dumb) = @_;
#   my $port_group; my @ret_ports;
#   
#   	my $new_port_string = "";
#	foreach my $card_group (split(/\;/,$port_string))
#	{
#		if ($card_group =~ /^ge\.(\d+)\./)
#		{
#		my $this_card = $1;
#		$card_group =~ s/^(ge\.\d+)\.//;
#		foreach my $port_group (split(/\,/,$card_group))
#		{
#			$new_port_string .= $this_card."\.".$port_group.",";
#		}
#	}
#		if ($card_group =~ /^tg\.(\d+)\./)
#		{
#		my $this_card = $1;
#		$card_group =~ s/^(tg\.\d+)\.//;
#		foreach my $port_group (split(/\,/,$card_group))
#		{
#			$new_port_string .= $this_card."\.".$port_group.",";
#		}
#	}
#		
#   }
#	$new_port_string =~ s/\,$//;
#	if ($new_port_string ne "")
#	{
#		$port_string = $new_port_string;
#	}	
#
#   foreach $port_group (split(/\,/,$port_string))
#   {
#      #Remove any padding spaces ie  0/ 0 to 0/0
#      $port_group =~ s/\s+//g;
#	  
#	  
#	  #Convert EXOS style to modular style
#      #ie 4:2-5 will be 4:2-4:5
#      #
#      $port_group =~ s/^(\d+)\:(\d+)-(\d+)$/$1\:$2\-$1\:$3/;
#	  
#	  #Convert EOS style to modular style
#      #ie 4.2-5 will be 4.2-4.5
#      #
#      $port_group =~ s/^(\d+)\.(\d+)-(\d+)$/$1\.$2\-$1\.$3/;
#
#	  #Convert stack style to modular style
#      #ie 4/2-5 will be 4/2-4/5
#      #
#      $port_group =~ s/^(\d+)\/(\d+)-(\d+)$/$1\/$2\-$1\/$3/;
#
#      #Convert single stack style to modular style
#      #ie  1-5,6,7,10-20 = 1/1-1/5,1/6,1/7,1/10-1/20
#
#      #Ranges:
#      #$port_group =~ s/^(\d+)-(\d+)$/1\/$1\-1\/$2/;
#
#      #Single ports:
#      #$port_group =~ s/^(\d+)$/1\/$1/;
#
#      if ($port_group =~ /^(\d+)\:(\d+)-(\d+)\:(\d+)$/)
#      {
#         #EXOS style 4:2-4:5 meaning 4/1,4/2,4/3,4/4,4/5
#         my $card = $1; my $sport = $2; my $lport = $4;
#
#         for(my $i=$sport;$i<$lport+1;$i++)
#         {
#            push(@ret_ports,"$card\/$i");
#			#print Dumper \@ret_ports;
#         }
#      }
#	  
#	  if ($port_group =~ /^(\d+)\.(\d+)-(\d+)\.(\d+)$/)
#      {   
#         #EOS style 4.2-4.5 meaning 4.1,4.2,4.3,4.4,4.5
#         my $card = $1; my $sport = $2; my $lport = $4;
#
#         for(my $i=$sport;$i<$lport+1;$i++)
#         {
#            push(@ret_ports,"$card\.$i");
#						
#         }
#      }
#	  
#	  
#	  if ($port_group =~ /^(\d+)\/(\d+)-(\d+)\/(\d+)$/)
#      {
#         #Modular style 4/2-4/5 meaning 4/1,4/2,4/3,4/4,4/5
#         my $card = $1; my $sport = $2; my $lport = $4;
#
#         for(my $i=$sport;$i<$lport+1;$i++)
#         {
#            push(@ret_ports,"$card\/$i");
#			print "11111\n";
#         }
#      }
#	  
#	  if ($port_group =~ /^(\d+)-(\d+)$/)
#      {
#         #Exos  style 1-5 meaning 1,2,3,4,5
#          my $sport = $1; my $lport = $2;
#
#         for(my $i=$sport;$i<$lport+1;$i++)
#         {
#            push(@ret_ports,"$i");
#			print "2222\n";
#         }
#      }
#	  
#	  
#	  
#      elsif ($port_group =~ /^(\d+)\/(\d+)\/(\d+)-(\d+)\/(\d+)\/(\d+)$/)
#      {
#         my $card = $1; my $st_port = $2; my $st_sport = $3; my $lst_port = $4; my $lst_sport = $5;
#         if ($st_port ne $lst_port)
#         {
#            #Modular sub-port style 1/41/1-1/42/4 meaning 1/41/1,1/41/2,1/41/3,1/41/4,1/42/1,1/42/2,1/42/3,1/42/4
#            #However, don't want to assume only 4 sub ports, so need to attempt calc and then compare against ALL_PORTS_ARRAY
#            #Will check for sub-ports 1 thru 128
#            for(my $i=$st_port;$i<$lst_port+1;$i++)
#            {
#               for(my $j=$st_sport;$j<129;$j++)
#               {
#                  if ($PORT_DATA{"$card\/$i\/$j"})
#                  {
#                     push(@ret_ports,"$card\/$i\/$j");
#                  }
#               }
#            }
#         }
#         else
#         {
#            #Modular sub-port style 4/2/1-4/2/4 meaning 4/2/1,4/2/2,4/2/3,4/2/4
#            for(my $i=$st_sport;$i<$lst_sport+1;$i++)
#            {
#               push(@ret_ports,"$card\:$st_port\/$i");
#			   print "3333\n";
#            }
#         }
#      }
#      elsif ($port_group =~ /^(\d+)\/ALL$/i)
#      {
#         #Stackable style 1/ALL, 2/ALL, etc.
#         my $card = $1;
#         #Because the ALL_PORT_ARRAY is required for this
#         #to work, it need to be populated first.
#         foreach $port (split(/\,/,$PORT_DATA{"ALL_PORTS_ARRAY"}))
#         {
#            my $this_card; my $this_port; my $this_sub_port;
#            ($this_card,$this_port,$this_sub_port) = split(/\//,$port);
#            if ($card == $this_card)
#            {
#               push(@ret_ports,$port);
#            }
#         }
#      }
#      elsif ($port_group =~ /^ALL$/i)
#      {
#         #Stackable style ALL
#         #Because the ALL_PORT_ARRAY is required for this
#         #to work, it need to be populated first.
#         foreach $port (split(/\,/,$PORT_DATA{"ALL_PORTS_ARRAY"}))
#         {
#            push(@ret_ports,$port);
#         }
#      }
#      elsif ($port_group =~ /^NONE$/i)
#      {
#         #Stackable style NONE, used for the VLAN 1.
#         #Do nothing.
#      }
#      else
#      {
#			#print Dumper \@ret_ports;
#			push(@ret_ports,$port_group);
#      }
#   }
#   #return(join(",",@ret_ports)); # for sending as string
#   return(@ret_ports);  #for sending as an array
#}


sub convert_to_full_vlan_list
{
   my $vlan_string; my $vlan;
   ($vlan_string,@dumb) = @_;
   my $vlan_group; my @ret_vlans;
   
   	
   foreach $vlan_group (split(/\,/,$vlan_string))
   {
      
	  if ($vlan_group =~ /^(\d+)-(\d+)$/)
      {
         #VLAN  1-5 meaning 1,2,3,4,5
          my $svlan = $1; my $lvlan = $2;

         for(my $i=$svlan;$i<$lvlan+1;$i++)
         {
            push(@ret_vlans,"$i");
		   }
      }
	  
	  else
      {
			#print Dumper \@ret_ports;
			push(@ret_vlans,$vlan_group);
      }
   }
   #return(join(",",@ret_ports)); # for sending as string
   return(@ret_vlans);  #for sending as an array
}


sub trim {

	my $line = shift;
   #Clean up line.
   chomp($line);
   $line =~ s/^\s+//;
   $line =~ s/\s+$//;
   $line =~ s/\r//;  
   $line =~ s/_x000D_//g;      ### This regex is added to remove line break from Excell Cell
   return $line;
 }
 
 sub clearVars()
{
    $vlanName = "";
    $vlanID = "";
    #$vlanType = "";
    #$vlanIPAddress = "";
    #$vlanSubnetMask = "";
    #$vrrpID = "";
    #$vrrpAddress = "";
    #$vrrpPriority = "";
    #$vrf = "";
    #$moreParms = 0;
    #$stgID = "";
    #$ospfEnable = "";
    #$ospfIntType = "";
    #$vlanIsid = "";
	#$protocol ="";
}