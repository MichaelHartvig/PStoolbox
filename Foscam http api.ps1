#set timezone via commandline

$CameraIP = "192.168.1.100:88" 
$Username = "admin" 
$Password = "yourpassword"
$timezone = "-3600" #UTC-1
$timeformat = "1" #0 = 12H / 1 = 24H

curl "http://$CameraIP/cgi-bin/CGIProxy.fcgi?cmd=setSystemTime&usr=$Username&pwd=$Password&timeSource=1&ntpServer=pool.ntp.org&timeZone=$timezone&timeFormat=$timeformat"

#reboot after set
curl "http://$CameraIP/cgi-bin/CGIProxy.fcgi?cmd=rebootSystem&usr=$Username&pwd=$Password"
