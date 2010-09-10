
=begin rdoc
*Revisions*
  | Initial File                              | Scott Shanks        | 9/10/2010|

*Module_Name*
  Bacnet

*Description*
  Bacnet methods - This is a wrapper for the bacnetconsole application created
by Chad Mrazek

http://126.4.1.113/twiki/bin/view/LmgEmbedded/MAT_BACnet_Client_Command_Line.

BACnet UDP Command Line Read / Write Tool
OPTIONS:
 /IP=, IP Address
 /P=, Send To Port
 /B=, Object Type
      ANALOG-INPUT
      ANALOG-OUTPUT
      ANALOG-VALUE
      BINARY-INPUT
      BINARY-OUTPUT
      BINARY-VALUE
      MULTISTATE-INPUT
      MULTISTATE-OUTPUT
      MULTISTATE-VALUE
 /I=, Instance Number
 /R=, Property ID
 /V=, Write Value, if present a write will occur
 /T=, Timeout in milliseconds

*Variables*

=end

module  Bacnet
  PATH_TO_BACNETCONSOLE = "C:\\Program Files\\Bacnet\\Console\\bacnetConsole.exe"

  def bacnet_get_object(instance_number, ip_address, property_id=PRESENT_VALUE,
                        object_type='ANALOG-VALUE', port=47808, timeout=1000)
    command = "#{PATH_TO_BACNETCONSOLE} /IP #{ip_address} /P #{port}
               /B #{object_type} /I #{instance_number} /R #{property_id},
               /T #{timeout}"
  end

  def bacnet_set_object(instance_number, ip_address, property_id=PRESENT_VALUE,
                        object_type='ANALOG-VALUE', port=47808, timeout=1000)
    command = "#{PATH_TO_BACNETCONSOLE} /IP #{ip_address} /P #{port}
               /B #{object_type} /I #{instance_number} /R #{property_id},
               /T #{timeout}"
  end

  def bacnet_parse(ws)
    @start_time = Time.now
    @total_rows = @ws.Range("A65536").End(XLUP).Row
    @row = 2
    while (@row <= @total_rows)
      @inner_row = @row
      register_value = query_modbus(@ws.Range("b#{@row}")['Value'])
      if @bit_position != nil then
        if register_value.to_i == 0 then
          @ws.Range("g#{@inner_row}")['Value'] = 0
        else
          s = "%.16b" % register_value.to_i.abs.to_s(2)
          @ws.Range("g#{@inner_row}")['Value'] = s[s.size-1-@bit_position.to_i].chr
        end
        @inner_row += 1
      else
        register_value.split.each do |s|
          @ws.Range("g#{@inner_row}")['Value'] = s
          @inner_row += 1
        end
      end
      @row += 1
      @wb.Save
    end
    @wb.Save
    @wb.Close
    @fin = Time.now
    p @fin
    @elapsed = (@fin - @start_time)
    puts " Elapsed time is seconds is: #{@elapsed}"
  end

end

