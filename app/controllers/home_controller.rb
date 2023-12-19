require 'roo'
require 'zip'

class HomeController < ApplicationController
  def index
  end

  def preview
    respond_to do |format|
      format.html
      format.pdf do
        pdf_content = WickedPdf.new.pdf_from_string(
          render_to_string(
            template: 'home/previewing.pdf.erb',
            layout: 'layouts/pdf.html.erb',
            formats: [:html]
          )
        )

        send_data(pdf_content, filename: 'Home_Index_Page.pdf', disposition: 'inline')
      end
    end
  end

  def print
    row = params[:row]
    row_number = params[:row_number]
    respond_to do |format|
      format.html
      format.pdf do
        render pdf: 'Admit Card - Cadet Sr.No #{row_number}',
               template: 'home/show',
               layout: 'layouts/pdf.html.erb',
               disposition: 'attachment',
               locals: { row: row }
      end
    end
  end

  def parse_excel
    zip_file_path = Rails.root.join('tmp', 'admit_cards.zip')
    File.delete(zip_file_path) if File.exist?(zip_file_path)
    pdf_paths = []  # Array to store paths for later combining into a single PDF

    if params[:excel_file].present? && params[:excel_file].content_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      excel_file = params[:excel_file].tempfile

      # Use Roo to read Excel data
      excel = Roo::Excelx.new(excel_file)

      # Assuming the data is in the first sheet
      sheet = excel.sheet(0)

      # Directory to store individual PDFs
      pdf_dir = Rails.root.join('tmp', 'pdfs')
      Dir.mkdir(pdf_dir) unless Dir.exist?(pdf_dir)

      # Iterate through all rows starting from row 2
      (2..sheet.last_row).each do |row_number|
        row = sheet.row(row_number)
        pdf_content = WickedPdf.new.pdf_from_string(
          render_to_string(
            template: 'home/show.pdf.erb',
            layout: 'layouts/pdf.html.erb',
            locals: { row: row },
            formats: [:html]
          )
        )

        # Save the PDF to a file
        pdf_path = File.join(pdf_dir, "Admit_card_#{row_number}.pdf")
        File.open(pdf_path, 'wb') { |file| file << pdf_content }
        pdf_paths << pdf_path
      end

      # Generate the zip file
      zip_file_path = Rails.root.join('tmp', 'admit_cards.zip')
      Zip::File.open(zip_file_path, Zip::File::CREATE) do |zipfile|
        pdf_paths.each do |pdf_path|
          zipfile.add(File.basename(pdf_path), pdf_path)
        end
      end

      # Send the zip file as a response
      send_file zip_file_path, type: 'application/zip', disposition: 'attachment', filename: 'admit_cards.zip'
    else
      flash[:error] = "Please upload a valid Excel file"
      redirect_to root_path
    end
  end
end
