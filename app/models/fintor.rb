require 'csv'
require 'simple_xlsx_reader'
require 'fileutils'

class FintorParser
  attr_accessor :input_2018, :input_2019, :input_2020, :output, :errors

  def initialize(input_2018, input_2019, input_2020, output)
    @input_2018 = input_2018
    @input_2019 = input_2019
    @input_2020 = input_2020
    @output     = output
  end

  def excel_check?
    self == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  end

  def output_file
    CSV.open(self.output, "wb") do |csv|
      csv << ["", "2020", "2019", "2018", "2017"]
    end
  end

  def processing
    # put dummy on unavailable excels
    if @input_2018.nil?
      doc_2018 = SimpleXlsxReader.open("./app/models/dummy.xlsx")
    else
      doc_2018 = SimpleXlsxReader.open(self.input_2018)
    end

    if @input_2019.nil?
      doc_2019 = SimpleXlsxReader.open("./app/models/dummy.xlsx")
    else
      doc_2019 = SimpleXlsxReader.open(self.input_2019)
    end

    if @input_2020.nil?
      doc_2020 = SimpleXlsxReader.open("./app/models/dummy.xlsx")
    else
      doc_2020 = SimpleXlsxReader.open(self.input_2020)
    end

    pg1_general         = doc_2019.sheets[1].rows.each(&:compact!) || doc_2020.sheets[1].rows.each(&:compact!) || doc_2018.sheets[1].rows.each(&:compact!)

    pg2_position_2018   = doc_2018.sheets[2].rows.each(&:compact!)
    pg3_profit_2018     = doc_2018.sheets[3].rows.each(&:compact!)
    pg5_cashflow_2018   = doc_2018.sheets[6].rows.each(&:compact!)

    pg2_position_2019   = doc_2019.sheets[2].rows.each(&:compact!)
    pg3_profit_2019     = doc_2019.sheets[3].rows.each(&:compact!)
    pg5_cashflow_2019   = doc_2019.sheets[6].rows.each(&:compact!)

    pg2_position_2020   = doc_2020.sheets[2].rows.each(&:compact!)
    pg3_profit_2020     = doc_2020.sheets[3].rows.each(&:compact!)
    pg5_cashflow_2020   = doc_2020.sheets[6].rows.each(&:compact!)

  # line item properties
    items = { :current_assets     => ["total current assets", "pg2_position", 55],
              :non_current_assets => ["total non-current assets", "pg2_position", 121],
              :total_assets       => ["total assets", "pg2_position", 122],
              :proceeds_tangible  => ["Proceeds from disposal of property, plant and equipment", "pg5_cashflow", 53],
              :bank_loans_short   => ["Current maturities of bank loans", "pg2_position", 165],
              :current_liabilities => ["Total current liabilities", "pg2_position", 187],
              :total_liabilities  => ["Total liabilities", "pg2_position", 231],
              :sales_revenue      => ["Sales and Revenue", "pg3_profit", 4],
              :selling_expenses   => ["Selling Expenses", "pg3_profit", 7],
              :general_expenses   => ["General and administrative expenses", "pg3_profit", 8],
              :operating_cash     => ["Total net cash flows received from (used in) operating activities", "pg5_cashflow", 101] }

    # rounding
    rounding            = pg1_general[24][1].split(" ").to_a.pop
    rounding_modifier   = { "Amount" => 1, "Million" => 1000000, "Thousand" => 1000 }
    rounding_multiplier = rounding_modifier.fetch(rounding).to_i

    # content extraction
    items.each do |k, v|
      item = []

      if v[1] == "pg2_position"
        item << [v[0], pg2_position_2020[v[2]][1], pg2_position_2019[v[2]][1], pg2_position_2018[v[2]][1], pg2_position_2018[v[2]][2]]
      elsif v[1] == "pg3_profit"
        item << [v[0], pg3_profit_2020[v[2]][1], pg3_profit_2019[v[2]][1], pg3_profit_2018[v[2]][1], pg3_profit_2018[v[2]][2]]
      elsif v[1] == "pg5_cashflow"
        item << [v[0], pg5_cashflow_2020[v[2]][1], pg5_cashflow_2019[v[2]][1], pg5_cashflow_2018[v[2]][1], pg5_cashflow_2018[v[2]][2]]
      else
        next
      end

      rows = item.flatten!

      # clean up label at content
      rows[1] = nil if rows[1].is_a? String
      rows[2] = nil if rows[2].is_a? String
      rows[3] = nil if rows[3].is_a? String

      # formatting numbers
      rows.map! do |row|
        if row.is_a? Float
          row.to_i * rounding_multiplier
        else
          row
        end
      end

      # output to console
      p rows

      # output to file
      CSV.open(self.output, "a+") do |csv|
        csv << rows
      end
    end
  end

  def output_name
  output.delete_prefix("public/output/")
  end
end
