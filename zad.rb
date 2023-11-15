require 'roo'
# require 'roo-xls'
class ExcelTable
  attr_accessor :data, :file_path, :headers

  def initialize(file_path)
    if file_path.end_with?('.xlsx')
      initialize_xlsx(file_path)
    elsif file_path.end_with?('.xls')
      initialize_xls(file_path)
    else
      raise ArgumentError, "Unsupported file format"
    end
  end

    def initialize_xlsx(file_path)
      spreadsheet = Roo::Excelx.new(file_path)
      @data = spreadsheet.parse(headers: true)
      @file_path = file_path
      @headers = @data.first.keys
      create_column_methods
    end

    def initialize_xls(file_path)
      spreadsheet = Roo::Spreadsheet.open(file_path, extension: :xlsx)
      @data = spreadsheet.parse(headers: true)
      @file_path = file_path
      @headers = @data.first.keys
      create_column_methods
    end

  def to_array
    return [] unless @data
    @data.to_a
  end


  def row(index)
    return nil unless @data && index.positive? && index <= @data.length
    @data[index - 1]
  end


  def each(&block)
    return unless @data

    @data.each do |row|
      handle_merged_cells(row, &block)
    end
  end


  def [](column_name)
    Column.new(@data, column_name)
  end

  def +(other_table)

    raise ArgumentError, "Can only add ExcelTable instances" unless other_table.is_a?(ExcelTable)

    raise ArgumentError, "Headers need to be the same" unless other_table.headers == @headers

    @data.concat(other_table.data[1..-1])

    self
  end

  def -(other_table)
    unless other_table.is_a?(ExcelTable)
      raise ArgumentError, "Can only subtract ExcelTable instances"
    end

    raise ArgumentError, "Headers need to be the same" unless other_table.headers == @headers

    subtracted_data = @data.reject do |row1|
      other_table.data.any? { |row2| row1.values == row2.values }
    end

    ExcelTable.new(file_path).tap { |table| table.data = subtracted_data }
  end



  def create_column_methods
    return unless @data

    headers = @data.first.keys

     headers.each do |column_name|
      method_name = column_name.gsub(' ', '')
      define_singleton_method(method_name) do
        Column.new(@data, column_name)
      end
    end
  end


  class Column
    attr_reader :data, :name

    def initialize(data, name)
      @data = data
      @name = name
    end

    def to_a
      @data.map { |row| row[@name] }
    end

    def [](index)
        return nil unless index.is_a?(Integer) && index.positive?

        to_a[index - 1]
      end


    def []=(index, value)
      if index.positive? && index <= to_a.length
        @data[index - 1][@name] = value
      end
    end

    def sum
      map(&:to_i).reduce(&:+)
    end

    def avg
      sum.to_f / @data.count { |i| i[@name].is_a?(Numeric) }
    end

    def extract_row(cell_value)
      @data.find { |row| row[@name] == cell_value }
    end

    def map(&block)
      to_a.map(&block)
    end

    def select(&block)
      to_a.select(&block)
    end

    def reduce(initial = nil, &block)
      to_a.reduce(initial, &block)
    end
  end

  def handle_merged_cells(row, &block)
    ignore_row = row.values.all?(&:nil?) || row.values.all? { |value| value.to_i.zero? }
    return if ignore_row

    row.each_with_index do |(header, value), index|
      next if value.nil?

      merged_cells = merged_cells_for_header(header)
      if merged_cells.any? { |range| range.include?(index) }
        next
      end
      yield value
    end
  end


  def merged_cells_for_header(header)
    []
  end
end

excel_file_path = 'test.xlsx'
excel_file_path2 = 'test2.xls'
excel_table = ExcelTable.new(excel_file_path)

table_array = excel_table.to_array

table_array.each { |row| puts row.values.join("\t") }

row_index = 1
row_data = excel_table.row(row_index)

if row_data
  puts "Red #{row_index}: #{row_data.values.join("\t")}"
else
  puts "Red #{row_index} Nije pronadjen."
end

puts "Sve celije: #{table_array.flatten.join("\t")}"

excel_table.each do |cell|
  puts "Celija: #{cell}"
end


first_column = excel_table['Prva Kolona']
puts "Kolone: 'Prva Kolona': #{first_column.to_a}"

value = first_column[2]
puts "Vrednost na indeksu 2: #{value}"

first_column[2] = 2556
puts "Modifikovana vrednost na indeksu 2: #{first_column[1]}"

table_array = excel_table.to_array
table_array.each { |row| puts row.values.join("\t") }

first_column = excel_table['Prva Kolona']

puts "Suma Prve Kolone: #{first_column.sum}"
puts "Srednja vrednost Prve Kolone: #{first_column.avg}"

cell_value = 2
row = first_column.extract_row(cell_value)
puts "Redovi sa #{cell_value} u Prvoj Koloni: #{row}"

mapped_values = first_column.map { |cell| cell.to_i * 2 }
puts "Mapirane vrednosti prve kolone: #{mapped_values}"

selected_values = first_column.select { |cell| cell.to_i > 3 }
puts "Vrednosti vece od 3 u prvoj koloni: #{selected_values}"

reduced_value = first_column.reduce(0) { |sum, cell| sum + cell.to_i }
puts "Redukovana suma Kolone1: #{reduced_value}"

table2 = ExcelTable.new(excel_file_path2)
puts "tabele:"
excel_table.to_array.each { |row| puts row.values.join("\t") }
table2.to_array.each { |row| puts row.values.join("\t") }

excel_table + table2
puts "tab1:"
excel_table.to_array.each { |row| puts row.values.join("\t") }

result_table = excel_table - table2
puts "rezultat"
result_table.to_array.each { |row| puts row.values.join("\t") }

puts "tabele:"
excel_table.to_array.each { |row| puts row.values.join("\t") }
table2.to_array.each { |row| puts row.values.join("\t") }
