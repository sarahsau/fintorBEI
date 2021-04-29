class Converter < ApplicationRecord

  require_relative 'fintor'

  def run_conversion
    ticker     = self[:ticker]
    file_2020  = self[:file_2020]
    file_2019  = self[:file_2019]
    file_2018  = self[:file_2018]
    output     = "public/output/#{ticker}_fintor_#{SecureRandom.alphanumeric}.csv"
    statement  = FintorParser.new(file_2020, file_2019, file_2018, output)

    statement.output_file
    statement.processing
    return statement.output_name
  end
end
