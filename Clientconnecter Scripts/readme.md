
# VBS Script usage

The vbs Scripts have to be run in a command line shell with administrative privileges in this format:

 cscript name_of_the_vbs_script

# DNS Service Record
in Bind9 the it has to look like this

_vlmcs._tcp     IN  SRV 0   0   1688    kmsserver.mydomain.

# installing an Office or Windows Product in Network containing KMS Server
if you've defined a Service Record pointing to KMS Server on your DNS Server
and you are installing a Corporate Edition of Windows, or an Office Product you don't need to register to KMS.
This Products are searching on installation after KMS Server itself via looking up for Service Record.
So this Products will automatically registered and activated.
