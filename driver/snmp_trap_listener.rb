#The following line has to be in snmptrad.conf:
#
# traphandle default c:/Ruby/bin/ruby.exe c:/LMG_Test/ruby/Framework/driver/snmp_trap_listener.rb
#
#
format_string = "%-60s %-60s\n"
trap_text = $stdin.read
trap_text_array = trap_text.split("\n")

hostname = trap_text_array[0]
ip_address = trap_text_array[1]
varbinds = trap_text_array[2..trap_text_array.size-1]
description = trap_text_array[5].split(" ")

printf format_string, "Trap received from #{hostname}!", description[1..description.size-1]
#puts "Trap From: #{hostname}"
#puts "IP Address is: #{ip_address.split(':')[1].split('[')[2].split(']')[0]}"
printf format_string +"\n", 'ip information', ip_address

varbinds.each do |line|
  oid = line.split(' ')[0]
  value = line.split(' ')[1..line.split(' ').size-1].to_s
  #puts "VARBIND: #{oid},#{value}"
  printf format_string, oid, value
end
puts "\n\n"


