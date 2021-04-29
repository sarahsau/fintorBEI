class ConvertersController < ApplicationController

  def index
    @converter = Converter.new
  end

  def show
    redirect_to root_path
  end

  def create
    @converter  = Converter.new(converter_params)

    if converter_params[:ticker].nil?
      flash.now[:alert] = "Error: no ticker entered"
    end

    result      = @converter.run_conversion
    file_path   = Rails.root.join('public', 'output', result)

      flash.now[:success] = "Ekstrasi data berhasil"
      stream_then_delete_statement(file_path)
  end

  def faq
  render 'converters/faq'
  end

  private

  def converter_params
    params.require(:converter).permit(:ticker, :file_2020, :file_2019, :file_2018)
  end

  def format_ok?(statement_params)
    if !statement_params.content_type.excel_check?
      flash.now[:alert] = "Error: wrong file type. Should be .xlsx"
    end
  end

  def stream_then_delete_statement(file_path)
    File.open(file_path, 'r') do |f|
      send_data(f.read, filename: "fintor_#{converter_params[:ticker]}.csv", type: 'text/csv')
    end

    File.delete(file_path)
  end
end
