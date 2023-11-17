require "ecr"
require "compress/zip"

module Xlsx
  class Column
    alias Value = UInt8 | UInt16 | UInt32 | UInt64 | Int8 | Int16 | Int32 | Int64 | Float32 | Float64 | Bool | String | Nil

    getter name : String?
    getter value : Value

    def initialize(@name : String?, @value : Value = nil)
      raise Exception.new("Can not use a type which can not be converted to String") unless @value.responds_to?(:to_s)
    end
  end

  class Row
    getter columns : Array(Column) = [] of Column

    def initialize(@columns : Array(Column))
    end
  end

  class Alphabet
    getter definitions : Array(String) = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]

    def name(index : Int32) : String
      if index <= 25
        definitions[index]
      else
        current_index = ((index - 25) / 25).to_i
        [definitions[current_index], definitions[index - 26]].join
      end
    end
  end

  class Converter
    def convert(value : String) : Array(Row)
      lines = value.split("\n")

      rows = [Row.new([] of Column)] of Row

      lines.first.split(",").each do |name|
        column = Column.new(name: name)
        rows.first.columns.push(column)
      end

      lines.shift

      lines.each do |line|
        columns = [] of Column
        crumbs = line.split(",")

        crumbs.each do |crumb|
          columns.push(Column.new(name: nil, value: crumb))
        end

        rows.push(Row.new(columns))
      end

      rows
    end
  end
end
