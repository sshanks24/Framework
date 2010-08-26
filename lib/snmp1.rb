
=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Module_Name*
  Snmp

*Description*
  Snmp methods - this is a wrapper for NetSNMP

*Variables*
  ip      = ip address
  com_str = community string
  mib     = management information base
  oid     = object id
  type    = value type
  val     = value
=end


#TODO snmp_set is broken
module  Snmp
  PATH_TO_SNMPTRAPD = "C:\\usr\\bin\\snmptrap.exe"
  PATH_TO_SNMPTRAPD_CONF = "C:\\usr\\etc\\snmp\\snmptrapd.conf"
  CONVERSION_THRESHOLD = 1
  #get SNMP value through system command
  def snmp_get(oid,version='2c',community_string=@community_string,ip=@test_site)
    command = "snmpget -v#{version} -c #{community_string} #{ip} #{oid}"
    puts "get command=#{command}"
    # Convert return oid to array and extract value from 4th element[3]
    snmp_data = `#{command}`.to_s.split(/ /)
    type = snmp_data[2]
    case type
    when /INTEGER:/ then return snmp_data[3].to_s.strip
    when /Gauge32:/ then return snmp_data[3].to_s.strip
    when /STRING:/ then return snmp_data[3..snmp_data.size].collect {|x| x + " "}.to_s.strip
    end
      return snmp_data[3..snmp_data.size].collect {|x| x + " "}.to_s.strip
  end

  #set SNMP value through system command
  def snmp_set(oid,type,value,version='2c',community_string=@community_string,ip=@test_site)
    case type
    when 'i' then command = "snmpset -v#{version} -c #{community_string} #{ip} #{oid} #{type} #{value.to_i}"
    when 'u' then command = "snmpset -v#{version} -c #{community_string} #{ip} #{oid} #{type} #{value.to_i}"
    when 's' then command = "snmpset -v#{version} -c #{community_string} #{ip} #{oid} #{type} #{value}"
    else command = "snmpset -v#{version} -c #{community_string} #{ip} #{oid} #{type} #{value}"
    end
    puts "set command=#{command}"
    `#{command}`
    sleep(3)
  end

  #Performs an snmpwalk, parses each line of output and returns as an array in
  #the format: OID, Value, Type, OID, Value, Type,...
  def snmp_walk(ip=@test_site,com_str=@community_string,version='2c',oid='.1')
    command = "snmpwalk -v #{version} -c #{com_str} #{ip} #{oid}"
    snmp_data = `#{command}`.to_s

    a = Array.new
    snmp_data.each_line do |line|
      if line =~ /No more variables/ then next; end;
      a << line.split('=')[0].strip # This should be the oid
      line.split('=')[1].each do |value|
        value.split(': ')[1].strip.each do |str| #This should be the value + any relavent units
          case value.split(':')[0].strip #This case statement parses the units according to type.
          when /string/i then a << str << ""
          when /oid/i then a << str << ""
          #This mess below probably deserves an explanation...
          #We are parsing the value + unit string by ' ', add the first piece to
          #the array we will return, and re-combine the remainder of what we
          #parsed to add to the array as one string.
          #There is probably a cleaner way of doing this... but it works for now
          #e.g. 320 .1 degrees celsius
          # a << 320 << .1 + ' ' + 'degrees' + ' ' + celsius + ' '
          when /integer/i then a << str.split(' ')[0] << str.split(' ')[1..str.split(' ').length-1].collect {|x| x + " "}.to_s
          when /gauge32/i then a << str.split(' ')[0] << str.split(' ')[1..str.split(' ').length-1].collect {|x| x + " "}.to_s
          else a << str << ""
          end
        end
        a << value.split(':')[0].strip  #This should be the type (e.g STRING)
      end
      end
    return a
  end

  def snmp_trap_listener(com_str=@community_string,path_to_trapd=PATH_TO_SNMPTRAPD,path_to_conf=PATH_TO_SNMPTRAPD_CONF)


    begin
    #Create backup file of snmptrapd.conf
    if File.exist?(path_to_conf)
      File.cp(path_to_conf, path_to_conf + '.bak')
    end
    #Build new configuration file

      # Start listening
    ensure
      #Make sure that we restore the original file

      #Clean up the old file
    end

  end

  #Runs the net-snmp snmptranslate command and parses the out returning the value
  #of max-access.
  def snmp_max_access(oid)
    command = 'snmptranslate -Td '<< oid
    snmp_data = `#{command}`
    snmp_data.each_line do |line|
      if line =~ /MAX-ACCESS/i
        return line.split(' ')[1].to_s
      end
    end
    "No Access Defined for #{oid}"
  end

  def snmp_get_type(oid)
    full_type = ''
    command = 'snmptranslate -Td -IR ' << oid.to_s
    snmp_data = `#{command}`
    snmp_data.each_line do |line|
      if line =~ /SYNTAX/i
        full_type = line.split(' ')[1].to_s
      end
    end
    type = case full_type #TODO Need to define the rest of the types...
    when /Integer/i then 'i'
    when /OCTET/ then 's'
    when /Unsigned32/ then 'u'
    else 's'
    end
    return type
  end

  def snmp_test_rw_oid(oid,value,expected,version='2c')
    command = "snmpget -v#{version} -c #{@community_string} #{@test_site} #{oid}"
    snmpget_result = `#{command}`

    # Discern the type, value, and units and translate into a form that net-snmp understands
    type = snmpget_result.split(' ')[2].to_s.strip.chop
    result_size = snmpget_result.split(' ').size #Will use this later to properly parse units
    puts "Type:#{type}"
    net_snmp_type = case type
    when /integer/i then
      units = snmpget_result.split(' ')[4..result_size].collect {|x| x + " "}.to_s.chop
      expected = expected.to_i
      net_snmp_type = 'i'
    when /string/i then
      units = ''
      net_snmp_type = 's'
    when /gauge32/i then
      units = snmpget_result.split(' ')[4..result_size].collect {|x| x + " "}.to_s.chop

      expected = expected.to_i
      net_snmp_type = 'u'
    else
      units = ''
      net_snmp_type = 's'
    end

    snmp_set(oid,net_snmp_type,value)
    new_value = snmp_get(oid)
    return new_value,process_test(value,new_value,net_snmp_type,units)

  end

  :private

  def process_test(expected,received,type,units)
    case units
    when /fahrenheit/i
      if received.to_i == expected.to_i or
         received.to_i-CONVERSION_THRESHOLD == expected.to_i or
         received.to_i+CONVERSION_THRESHOLD == expected.to_i
       return "PASS!"
      else
        return "FAIL: #{received} is not within #{CONVERSION_THRESHOLD} degree of #{expected}"
      end
    end
    case type
    when 'i'
      #Parse resolved mib objects for integer value (e.g. enabled(1), warning(4)...)
      if received.include? '(' then
        puts "Before: #{received}"
        received = received.split('(')[1].chop
        puts "After: #{received}"
      end
      if received.to_i == expected.to_i
        return "PASS!"
      else
        return "FAIL: #{received.to_i} != #{expected.to_i}"
      end
    when 'u'
      if received.to_i == expected.to_i
        return "PASS!"
      else
        return "FAIL: #{received} != #{expected}"
      end
    when 's'
      if received.to_s == expected.to_s
        return "PASS!"
      else
        return "FAIL: #{received} != #{expected}"
      end
    end
  end

end

