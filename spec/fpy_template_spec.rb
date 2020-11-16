RSpec.describe FpyTemplate do
  it "has a version number" do
    expect(FpyTemplate::VERSION).not_to be nil
  end

  it "check input file exists" do
    expect(File.exist?("excel_files/excel.xlsx")).to eq(true)
  end
  it "create output" do
    FpyTemplate::Parse.transform('excel_files/excel.xlsx')
    expect(File.exist?("output.xlsx")).to eq(true)
  end
  it "Match the contents for excel file " do 
    expected = RubyXL::Parser.parse('excel_files/output.xlsx')
    output = RubyXL::Parser.parse('output.xlsx')
    expected[0].each_with_index do |expected_val, i|
        rows= expected[0][0].size
        j=0
        while j<rows do
            expect(expected_val[j]&.value).to eq(output[0][i][j]&.value)
            j+=1
        end
    end
  end
  it "Clean up files" do 
    f = 'output.xlsx'
    File.delete(f) if File.exists? f
    expect(File.exist?(f)).to eq false
  end
end
