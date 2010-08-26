=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Module_Name*
  Xls

*Description*
  excel methods

*Variables*
  xls = instance of excel
    ss = spreadsheet
    wb = workbook
    ws = worksheet
  num = worksheet number
  s_s = workbook file name
  save_as = workbook file name to be saved as
     
=end

require 'win32ole'

module Xls
  #
  #  - createand return new instance of excel
  def new_xls(s_s,num) #wb name and sheet number
    ss = WIN32OLE::new('excel.Application')
    wb = ss.Workbooks.Open(s_s)
    ws = wb.Worksheets(num)
    ss.visible = true # For debug
    xls = [ss,wb,ws]
  end


  #
  #  - connect to an existing instance of excel and open new workbook
  #  - return an instannce of excel
  def conn_open_xls(s_s,num) #wb name and sheet number
    ss = WIN32OLE.connect('excel.Application')
    wb = ss.Workbooks.open(s_s)
    ws = wb.Worksheets(num)
    ss.visible = true # For debug
    xls = [ss,wb,ws]
  end


  #
  #  - connect to an existing instance of excel and open the existing workbook
  def conn_act_xls  #sheet number is currently hardcoded
    ss = WIN32OLE.connect('excel.Application')
    wb = ss.ActiveWorkbook
    ws = wb.Worksheets(1)
    ss.visible = true # For debug
  end
  

  #
  #  - save an existing workbook as another file name
  def save_as_xls(s_s,save_as)
    sleep 1
    s_s[2].saveas(save_as)
  end


  #
  #  - close an existing workbook
  def close_xls(s_s) # sheet number
    s_s[1].close
    s_s.quit # session will remain active until the ruby script has completed
  end  


  #
  #  - activate one of multiple open workbooks
  def act_xls(s_s)
    ss.workbooks(s_s).activate
  end  
    
end

