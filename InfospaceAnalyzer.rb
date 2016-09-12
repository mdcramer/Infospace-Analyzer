class String
  def cap_words
    self.split(' ').map {|w| w.capitalize }.join(' ')
  end
  def decrement_excel_pointers
    self.gsub(/[A-Z]+\d+/) {|match| match.scan(/\D+/).join + (match.scan(/\d+/).join.to_i - 1).to_s}
  end
end

class Fixnum
  def days
    self * 60 * 60 * 24
  end
end

require 'date'
require 'csv'
# browsers = ["Chrome", "FF", "IE", "IE9", "Other"]
countries = ['United States', 'Australia', 'Canada', 'France', 'Germany', 'Italy', 'Netherlands', 'Spain', 'United Kingdom']

#Load W3i installation data into w3i_installs hash
puts 'Importing W3i install data...'
w3i_installs = {}
CSV.foreach('W3i Tracking.CSV') do |line| #pull in W3i install numbers from Tracking.txt - Browser, Country, Date, Views, Accepts
  date = line[2].split(/\//) #get date and split at "/"
  if date[0].to_i != 0
    date = Time.gm(date[2].to_i, date[0].to_i, date[1].to_i) #convert date to GMT
    browser = case line[0].strip
                when /chrome/i then
                  'Chrome'
                when /firefox/i then
                  'FF'
                when /msie 9/i then
                  'IE9'
                when /msie/i then
                  'IE'
                else
                  'Other'
              end
    browserhash = ( w3i_installs[date] != nil ? w3i_installs[date] : {} )
    installhash = ( browserhash[browser] != nil ? browserhash[browser] : {} )
    if installhash[line[1].strip.cap_words] == nil
      installhash[line[1].strip.cap_words] = {:views => line[3].to_i, :accepts => line[4].to_i}
    else
      installhash[line[1].strip.cap_words] = installhash[line[1].strip.cap_words].merge({:views => line[3].to_i, :accepts => line[4].to_i}) {|key, v1, v2| v1 + v2} #add the new values to the old values
    end
    if installhash[:total] == nil
      installhash[:total] = {:views => line[3].to_i, :accepts => line[4].to_i}
    else
      installhash[:total] = installhash[:total].merge({:views => line[3].to_i, :accepts => line[4].to_i}) {|key, v1, v2| v1 + v2}
    end
    browserhash[browser] = installhash
    w3i_installs[date] = browserhash
  end
end
w3i_installs.keys.each { |date|
  w3i_installs[date].keys.each { |browser|
    countries.each { |country|
      if country == 'United States'
        w3i_installs[date][browser][:eu] = {}
      else
        if w3i_installs[date][browser][country] == nil
          w3i_installs[date][browser][country] = {:views => 0, :accepts => 0} #set to zero if there is no data
        end
        w3i_installs[date][browser][:eu] = w3i_installs[date][browser][:eu].merge({:views => w3i_installs[date][browser][country][:views], :accepts => w3i_installs[date][browser][country][:accepts]}) { |key, v1, v2| v1 + v2 }
      end
    }
    w3i_installs[date][browser][:rw] = {}
    if w3i_installs[date][browser]['United States'] == nil
      w3i_installs[date][browser]['United States'] = {:views => 0, :accepts => 0}
    end
    w3i_installs[date][browser][:rw][:views] = w3i_installs[date][browser][:total][:views] - w3i_installs[date][browser]['United States'][:views] - w3i_installs[date][browser][:eu][:views]
    w3i_installs[date][browser][:rw][:accepts] = w3i_installs[date][browser][:total][:accepts] - w3i_installs[date][browser]['United States'][:accepts] - w3i_installs[date][browser][:eu][:accepts]
  }
}

#pull in Infospace data and insert into infospace_hash
puts 'Importing Infospace data...'
infospace_hash = {} #["Affiliate", {"Day"=>{:clicks=>0.0, :searches=>0.0, :revenue=>0.0, :net=>0.0, :vintage=>"none"}}]
infospacearray = File.readlines('Search Performance Metrics.CSV') #pull in new data from InfoSpace
infospaceaffiliates = [nil, nil, 'PredictAd', 'WOT', nil, 'W3i IE9 US vintages I & J', 'W3i IE9 EU vintages I & J', 'W3i Chrome vintage A', 'W3i Firefox US vintages A & B', 'W3i Firefox EU vintages A & B', 'W3i Firefox US vintages A thru D', 'W3i Firefox EU vintages A thru D', 'W3i Rest of World', 'W3i IE US vintage F', 'W3i IE US vintage G', 'W3i IE US vintage H', 'W3i IE EU vintage H', 'W3i IE US vintages I & J', 'W3i IE EU vintages I & J', 'W3i IE US vintage E', 'W3i IE EU vintage E']
0.upto(infospacearray.length - 1) do |line|
  infospacearray[line].scan(/".*?"/) { |numinquotes| infospacearray[line] = infospacearray[line].sub(numinquotes, numinquotes.scan(/\d/).join) } #convert numbers with commas in quotes to just numbers with no quotes
  infospacearray[line] = infospacearray[line].scan(/\w|,|\/|\./).join.split(/,/) #remove extra characters and split at commas
  if infospacearray[line].length != 0
    if infospacearray[line][0].scan(/\d/) != [] #if there's a date, convert it to GMT
      date_gmt = infospacearray[line][0].split(/\/|' '/) #get date and split at "/"
      date_gmt = Time.gm(date_gmt[2], date_gmt[0], date_gmt[1]) #convert date to GMT
      infospacearray[line][0] = date_gmt
    end
    infospace_hash[infospacearray[line][2]] != nil ? innerhash = infospace_hash[infospacearray[line][2]] : innerhash = Hash.new({:clicks=>0.0, :searches=>0.0, :revenue=>0.0, :net=>0.0, :vintage=> 'none'})
    innerhash[infospacearray[line][0]] = {:clicks => infospacearray[line][3].to_f, :searches => infospacearray[line][5].to_f, :revenue => infospacearray[line][6].to_f, :net => infospacearray[line][6].to_f * 0.65}
    innerhash[infospacearray[line][0]][:vintage] = (infospacearray[line][2].scan(/\d/) != [] ? infospaceaffiliates[infospacearray[line][2].scan(/\d/).join.to_i] : 'none')
    infospace_hash[infospacearray[line][2]] = innerhash
  end
end

require 'win32ole'
excel = WIN32OLE.new('Excel.Application')
excel.visible = true
workbook = excel.workbooks.open('C:\Users\Mark\Documents\Surf Canyon\Accounting\Payables\Refinement Partner Clicks.xlsx')
w3i_worksheet = workbook.worksheets('W3i')
w3i_worksheet.Activate
days_to_load = 0

#Insert rows in installation tables at bottom of worksheet and insert data from W3i
puts 'Inserting W3i installation data into Excel...'
['Chrome', 'FF', 'IE'].each do |browser|
  data = w3i_worksheet.UsedRange.Value
  found = false
  row = 0
  until found
    # puts "#{browser} checking row #{row}"
    if data[row][1] == browser + ' Installs'
      row_with_offset = row + 3 #table begins a couple rows after the title
      rows_to_insert = (Date.today - Date.civil(data[row+2][0].to_a[5], data[row+2][0].to_a[4], data[row+2][0].to_a[3])).to_i - 1
      days_to_load = rows_to_insert
      if rows_to_insert > 0
        w3i_worksheet.Rows("#{row_with_offset}:#{row_with_offset + rows_to_insert - 1}").Insert
      end
      (3 + rows_to_insert - 1).downto(3) do |fill_row|
        w3i_worksheet.Cells(row + fill_row, 1).Value = (Date.parse(w3i_worksheet.Cells(row + fill_row + 1, 1).Value.to_s.split[0]) + 1).strftime('%F')
        date_gmt = w3i_worksheet.Cells(row + fill_row, 1).Text.split(/\//) #get date and split at "/"
        date_gmt = Time.gm(date_gmt[2].to_i + 2000, date_gmt[0], date_gmt[1]) #convert date to GMT
        0.upto(countries.length - 1) do |country|
          # add logic to handle case the no downloads
          if w3i_installs[date_gmt] == nil
            w3i_installs[date_gmt] = {}
          end
          if w3i_installs[date_gmt][browser] == nil
            w3i_installs[date_gmt][browser] = {}
          end
          if w3i_installs[date_gmt][browser][countries[country]] == nil
            w3i_installs[date_gmt][browser][countries[country]] = {:views => 0, :accepts => 0}
          end
          w3i_worksheet.Cells(row + fill_row, country + 2).Value = w3i_installs[date_gmt][browser][countries[country]][:accepts]
        end
        # add logic to handle case with no downloads
        if w3i_installs[date_gmt][browser][:total] == nil
          w3i_installs[date_gmt][browser][:total] = {:views => 0, :accepts => 0}
        end
        if w3i_installs[date_gmt][browser][:rw] == nil
          w3i_installs[date_gmt][browser][:rw] = {:views => 0, :accepts => 0}
        end
        if w3i_installs[date_gmt][browser][:eu] == nil
          w3i_installs[date_gmt][browser][:eu] = {:views => 0, :accepts => 0}
        end
        w3i_worksheet.Cells(row + fill_row, 11).Value = w3i_installs[date_gmt][browser][:total][:accepts]
        w3i_worksheet.Cells(row + fill_row, 12).Value = w3i_installs[date_gmt][browser][:rw][:accepts]
        w3i_worksheet.Cells(row + fill_row, 13).Value = w3i_installs[date_gmt][browser][:eu][:accepts]
        if browser == 'IE' #add IE9 data to right of IE data
          if w3i_installs[date_gmt]['IE9'] != nil
            if w3i_installs[date_gmt]['IE9']['United States'] != nil
              w3i_worksheet.Cells(row + fill_row, 14).Value = w3i_installs[date_gmt]['IE9']['United States'][:accepts]
            end
            if w3i_installs[date_gmt]['IE9'][:eu] != nil
              w3i_worksheet.Cells(row + fill_row, 15).Value = w3i_installs[date_gmt]['IE9'][:eu][:accepts]
            end
          end
        end
      end
      found = true
    end
    row += 1
  end
end

#Mapping of W3i vintages to Infospace channels. Chrome vintage A is split using MaxMind IP DB.
vintages = {
            #"IE7&8 US vintages A,B,C,D,F,G,H combined" => {:infospace_channel => "surfcanyon.01", :insertion_row => 24},
            #"IE7&8 US vintage E" => nil,
            #"IE7&8 US vintage I" => nil,
            #"IE7&8 US vintage J" => nil,
            #"IE7&8 US vintage K" => {:infospace_channel => "surfcanyon.05", :insertion_row => 8},
            #"IE EU vintages A,B,C,D,E,H combined" => {:infospace_channel => "surfcanyoneurope.01", :insertion_row => 24},
            #"IE7&8 EU vintage I" => nil,
            #"IE7&8 EU vintage J" => nil,
            #"IE7&8 EU vintage K" => {:infospace_channel => "surfcanyoneurope.05", :insertion_row => 8},
            #
            #"IE9 US vintage J" => {:infospace_channel => "surfcanyon.02", :insertion_row => 8},
            #"IE9 US vintage K" => {:infospace_channel => "surfcanyon.06", :insertion_row => 8},
            #"IE9 US vintage L" => {:infospace_channel => "surfcanyon.11", :insertion_row => 8},
            #"IE9 EU vintage K" => {:infospace_channel => "surfcanyoneurope.06", :insertion_row => 8},
            #"IE9 EU vintage L" => {:infospace_channel => "surfcanyoneurope.11", :insertion_row => 8},

            'IE US vintages A-O' => {:infospace_channel => 'surfcanyon.01', :insertion_row => 8},
            #"IE US vintage N" => {:infospace_channel => "surfcanyon.15", :insertion_row => 8},
            #"IE US vintage O" => {:infospace_channel => "surfcanyon.18", :insertion_row => 8},
            'IE EU vintages A-O' => {:infospace_channel => 'surfcanyoneurope.01', :insertion_row => 8},

            'Firefox US vintages A-O' => {:infospace_channel => 'surfcanyon.02', :insertion_row => 8},
            #"Firefox US vintage K" => {:infospace_channel => "surfcanyon.07", :insertion_row => 8},
            'Firefox EU vintages A-O' => {:infospace_channel => 'surfcanyoneurope.02', :insertion_row => 8},
            #"Firefox EU vintage K" => {:infospace_channel => "surfcanyoneurope.07", :insertion_row => 8},
            #"Firefox US vintage O" => {:infospace_channel => "surfcanyon.16", :insertion_row => 8},
            #"Firefox EU vintage O" => {:infospace_channel => "surfcanyoneurope.16", :insertion_row => 8},

            'Chrome vintages A-O' => nil,
            'Chrome US vintages A-O' => {:infospace_channel => 'surfcanyon.03', :insertion_row => 24},
            'Chrome EU vintages A-O' => {:infospace_channel => 'surfcanyoneurope.03', :insertion_row => 40},
            #"Chrome vintage K" => nil,
            #"Chrome US vintage K" => {:infospace_channel => "surfcanyon.08", :insertion_row => 24},
            #"Chrome EU vintage K" => {:infospace_channel => "surfcanyoneurope.08", :insertion_row => 40},
            #"Chrome US vintage O" => {:infospace_channel => "surfcanyon.17", :insertion_row => 8},
            #"Chrome EU vintage O" => {:infospace_channel => "surfcanyoneurope.17", :insertion_row => 8}
}

#Insert blank rows in Excel for each vintage
puts 'Inserting blank rows into Excel:'
vintages.keys.each do |vintage|
  data = w3i_worksheet.UsedRange.Value
  found = false
  row = 0

  #locate the row that matches the text of the vintage
  puts "--> #{vintage}..."
  until found || row == data.length - 1
    sleep(0.002) #slow down insertion by x seconds to let Excel catch up
    found = data[row][1] == vintage
    row += 1
  end

  if found == true
    row_with_offset = row + 2 #table begins a row after the title but start inserting below top row of table
    last_date_in_excel = data[row][0].to_a
    rows_to_insert = (Date.today - Date.civil(last_date_in_excel[5], last_date_in_excel[4], last_date_in_excel[3])).to_i - 1
    if rows_to_insert > 0
      w3i_worksheet.Rows("#{row_with_offset}:#{row_with_offset + rows_to_insert - 1}").Insert
    end
  end
  sleep(2) #slow down insertion by 2s to let Excel catch up
end

#copy functions and data into inserted rows
puts 'Inserting Infospace data into Excel:'
vintages.keys.each do |vintage|
  data = w3i_worksheet.UsedRange.Value
  found = false
  row = 0
  print "--> #{vintage}... "

  #locate the row that matches the text of the vintage
  until found || row == data.length - 1
    if data[row][1] == vintage
      found = true
      puts 'found'
    end
    #Special check for the Chrome vintage A data
    if (data[row][1] == 'Chrome vintages A-O') && ((vintage == 'Chrome US vintages A-O') || (vintage == 'Chrome EU vintages A-O'))
      found = true
      puts 'found for special processing...'
    end
    #Special check for the Chrome vintage K data
    #if (data[row][1] == "Chrome vintage K") && ((vintage == "Chrome US vintage K") || (vintage == "Chrome EU vintage K"))
    #  found = true
    #  puts "--> #{vintage}... special processing..."
    #end
    row += 1
  end

  if found == true
    (3 + days_to_load - 2).downto(1) do |fill_row|

      #increment date in column 1
      w3i_worksheet.Cells(row + fill_row, 1).Value = (Date.parse(w3i_worksheet.Cells(row + fill_row + 1,1).Value.to_s.split[0]) + 1).strftime('%F')

      #copy formulas up one row from columns 2 to 31
      2.upto(47) do |column|
        if w3i_worksheet.Cells(row + fill_row + 1, column).Formula[0] == '=' #check to see if cell is a formula
          w3i_worksheet.Cells(row + fill_row, column).Formula = w3i_worksheet.Cells(row + fill_row + 1, column).Formula.decrement_excel_pointers
        #else w3i_worksheet.Cells(row + fill_row, column).Formula = ""
        end
      end

      #insert Infospace data from infospace_hash
      if vintages[vintage] != nil
        if vintages[vintage][:infospace_channel] != nil #determine if there is Infospace data to insert
          date_gmt = w3i_worksheet.Cells(row + fill_row, 1).Text.split(/\//) #get date and split at "/"
          date_gmt = Time.gm(date_gmt[2].to_i + 2000, date_gmt[0], date_gmt[1]) #convert date to GMT
          column = vintages[vintage][:insertion_row]
          if w3i_worksheet.Cells(row + fill_row + 1, column).Formula[0] != '=' #check to see if cell is not a formula
            w3i_worksheet.Cells(row + fill_row, column).Value = infospace_hash[vintages[vintage][:infospace_channel]][date_gmt][:searches]
            w3i_worksheet.Cells(row + fill_row, column + 2).Value = infospace_hash[vintages[vintage][:infospace_channel]][date_gmt][:clicks]
            w3i_worksheet.Cells(row + fill_row, column + 4).Value = infospace_hash[vintages[vintage][:infospace_channel]][date_gmt][:revenue]
          else #sub in Infospace data in first number if the cell is a formula
            w3i_worksheet.Cells(row + fill_row, column).Formula = w3i_worksheet.Cells(row + fill_row, column).Formula.sub(/\d+/, infospace_hash[vintages[vintage][:infospace_channel]][date_gmt][:searches].to_s)
            w3i_worksheet.Cells(row + fill_row, column + 2).Formula = w3i_worksheet.Cells(row + fill_row, column + 2).Formula.sub(/\d+/, infospace_hash[vintages[vintage][:infospace_channel]][date_gmt][:clicks].to_s)
            w3i_worksheet.Cells(row + fill_row, column + 4).Formula = w3i_worksheet.Cells(row + fill_row, column + 4).Formula.sub(/\d+/, infospace_hash[vintages[vintage][:infospace_channel]][date_gmt][:revenue].to_s)
          end
        end
      end
    end
  end
end

puts 'Completed and saving...'
workbook.saved = true #don't prompt to save Excel file
workbook.Save
excel.ActiveWorkbook.Close(0)
excel.Quit()
puts 'Success!'
