class OPTRecord
	attr_accessor :pieces

	def initialize(line)
		@pieces = line.split(",")
	end

	def id
		return pieces[0]
	end

	def id=(value)
		pieces[0] = value
	end

	def volume
		return pieces[1]
	end

	def volume=(value)
		pieces[1] = value
	end

	def path
		return pieces[2]
	end

	def path=(value)
		pieces[2] = value
	end

	def is_first_page
		return pieces[3]
	end

	def is_first_page=(value)
		pieces[3] = value
	end

	def pages
		if pieces[6].strip.empty?
			return 0
		else
			return pieces[6].to_i
		end
	end

	def pages=(value)
		pieces[6] = value
	end

	def to_line
		return @pieces.join(",")
	end
end

# DOC-000000001_0001,,IMAGE\000\000\001\DOC-000000001_0001.tif,Y,,,2
# DOC-000000001_0002,,IMAGE\000\000\001\DOC-000000001_0002.tif,,,,

class OPT
	def self.each(file_path,&block)
		File.open(file_path,"r:utf-8") do |file|
			while line = file.gets
				line.chomp!
				record = OPTRecord.new(line)
				yield record
			end
		end
	end

	def self.transpose_each(input_file_path,output_file_path,&block)
		if output_file_path == input_file_path
			raise "input_file_path and output_file_path must be different locations"
		end
		File.open(output_file_path,"w:utf-8") do |output_file|
			File.open(input_file_path,"r:utf-8") do |file|
				while line = file.gets
					line.chomp!
					record = OPTRecord.new(line)
					yield(record)
					output_file.puts(record.to_line)
				end
			end
		end
	end
end