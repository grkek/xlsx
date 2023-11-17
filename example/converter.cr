require "../src/xlsx"

alphabet = Xlsx::Alphabet.new

# CSV provided in as a string
#
#
# Modify only the input line so none other will be altered
rows = Xlsx::Converter.new.convert("wow,woah,wow,haha,do,i,want,to,know,how,to,handle,all,of,this,information,yet,i,do,not,recognize,what,i,can,do,nor,do,i,know,how,to,handle,any,of,this,might,be,enough,to,break,this,little,silly,project,lets,see\n1,2,3,4")

sheet = ECR.render("#{__DIR__}/../template/xl/worksheets/sheet.xml.ecr")
shared_strings = ECR.render("#{__DIR__}/../template/xl/sharedStrings.xml.ecr")

Dir.mkdir("bin") unless Dir.exists?("bin")

File.write("#{__DIR__}/../template/xl/worksheets/sheet.xml", sheet)
File.write("#{__DIR__}/../template/xl/sharedStrings.xml", shared_strings)

File.open("./bin/template.xlsx", "w") do |file|
  Compress::Zip::Writer.open(file) do |zip|
    zip.add_dir("_rels")

    zip.add_dir("xl")
    zip.add_dir("xl/_rels")
    zip.add_dir("xl/drawings")
    zip.add_dir("xl/theme")
    zip.add_dir("xl/worksheets")
    zip.add_dir("xl/worksheets/_rels")

    zip.add("_rels/.rels", File.open("#{__DIR__}/../template/_rels/.rels"))

    zip.add("xl/_rels/workbook.xml.rels", File.open("#{__DIR__}/../template/xl/_rels/workbook.xml.rels"))

    zip.add("xl/drawings/drawing.xml", File.open("#{__DIR__}/../template/xl/drawings/drawing.xml"))

    zip.add("xl/theme/theme.xml", File.open("#{__DIR__}/../template/xl/theme/theme.xml"))

    zip.add("xl/worksheets/_rels/sheet.xml.rels", File.open("#{__DIR__}/../template/xl/worksheets/_rels/sheet.xml.rels"))
    zip.add("xl/worksheets/sheet.xml", File.open("#{__DIR__}/../template/xl/worksheets/sheet.xml"))

    zip.add("xl/sharedStrings.xml", File.open("#{__DIR__}/../template/xl/sharedStrings.xml"))
    zip.add("xl/styles.xml", File.open("#{__DIR__}/../template/xl/styles.xml"))
    zip.add("xl/workbook.xml", File.open("#{__DIR__}/../template/xl/workbook.xml"))

    zip.add("[Content_Types].xml", File.open("#{__DIR__}/../template/[Content_Types].xml"))
  end
end

File.delete("#{__DIR__}/../template/xl/worksheets/sheet.xml")
File.delete("#{__DIR__}/../template/xl/sharedStrings.xml")

puts "***ALL DONE***"