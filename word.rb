require 'win32ole'
require 'yaml'

class DocTable
	def initialize
		@word = WIN32OLE.new('Word.Application')
	end

	def count_temp temp_file
		@temp = @word.Documents.open(temp_file, 'ReadOnly' => true)
		
		all = {}
		# p @temp.tables.count
		table_cnt=1
		@temp.tables.each{|tb|
			# p tb.range.rows.count
			# p tb.range.columns.count

			(1..tb.rows.count).each{|row|
				(1..tb.columns.count).each{|col|
					begin
						# puts "#{row}x#{col} #{tb.cell(row, col).range.text}"
						if /\[(.+)\]/=~tb.cell(row, col).range.text
							name = $1
							# puts "#{row}x#{col} #{name}"
							all[name] = [table_cnt, row, col]
						end
					rescue Exception => e
						# puts "#{row}x#{col} NULL!}"
					end
				}
			}
			table_cnt +=1
		}
		@temp.close

		open("detect.yml","w") do |f|
			YAML.dump({worddoc: {table: all}}, f)
		end
	end
	
	def get_content doc_file
		detect = YAML::load_file('detect.yml')

		@doc = @word.Documents.open(doc_file, 'ReadOnly' => true)
		
		result = {}
		detect[:worddoc][:table].each{|name, value|
			# p name
			# p value
			result[name] = trim(@doc.tables(value[0]).cell(value[1], value[2]).range.text)
		}

		@doc.close
		result
	end
	
	def close
		@word.quit
	end
	
	protected
	def trim val
		val.gsub(/[\s\r\a\?]/, '').gsub(/[,]/, '，')
	end
end

def dump result
	result.each{|k, v|
		puts "#{k}: #{v}"
	}
end

if __FILE__==$0
	#read template
	# doc = DocTable.new
	# doc.count_temp("F:\\82_lwang\\CntWord\\CntWord\\temp.doc")
	# doc.close

	#write result
	doc = DocTable.new
	p doc.get_content("F:\\82_lwang\\CntWord\\CntWord\\论文\\孟蕊.doc")
	doc.close
end


