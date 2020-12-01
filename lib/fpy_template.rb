require "fpy_template/version"
require "fpy_template/helper"
# require "byebug"
require "rubyXL"
require 'rubyXL/convenience_methods'

module FpyTemplate
  class MissingColumns < StandardError; end
  class PositionNotDefined < StandardError; end
  class Parse < Helper
    
    def self.transform(file_path)
      input = RubyXL::Parser.parse(file_path)
      output = duplicate(input)
 
      verify_cols(output[0], ["identifier", "username", "email", "position", "area"])
      clean_data_for_column(output[0], ["identifier", "username", "email", "position"])
      downcase_cols(output[0], ["identifier", "username", "email"])
      verify_values_of_columns_exist(output[0],"position", ["Gerente General", "Gerente de Operaciones", "Bodeguero" , "Jefe de Obra", "Prevencionista", "Ayudante profesional", "Administrador de obra", "Profesional de obra","Administrativo","Jefe Administrativo","Capataz","Obra","Jefe de Bodega"])
      create_autoevaluation_and_ponderation_autoevaluation(output[0])
      create_column_and_return_index(output[0],"ponderation_ascendente")
      
      add_asendentees(output[0],["Administrador de Obra"],["Jefe de Obra","Ayudante Profesional","Jefe Administrativo","Jefe de Bodega","Bodeguero","Administrativo","Profesional de Obra","Prevencionista","Capataz"],"ascendente")
      add_asendentees(output[0],["Jefe de obra","Ayudante profesional","Jefe administrativo","Jefe de bodega","Administrativo","Bodeguero","Prevencionista","Profesional de Obra"],["Capataz"],"ascendente")
      # byebug
      create_column_and_return_index(output[0],"ponderation_gerenete")
      add_asendentees(output[0],["Administrador de Obra"],["Gerente General","Gerente de operaciones"],"gerente")
      create_column_and_return_index(output[0],"ponderation_descendente")
      add_asendentees(output[0],["Ayudante profesional", "Jefe de obra","Prevencionista","Bodeguero","Jefe de bodega","Obra","Capataz","Administrativo","Jefe administrativo","Profesional de Obra"],["Administrador de Obra"],"descendente")
      add_asendentees(output[0],["Capataz"],["Profesional de obra","Ayudante profesional","Administrativo","Jefe de bodega","Bodeguero","Prevencionusta","Jefe de obra","Jefe administrativo"],"descendente")
      create_column_and_return_index(output[0],"ponderation_equipo")
      add_asendentees(output[0],["Profesional de obra","Ayudante profesional","Administrativo","Jefe de bodega","Bodeguero","Prevencionista","Jefe de obra","Jefe administrativo"],["Profesional de obra","Ayudante profesional","Administrativo","Jefe de bodega","Bodeguero","Prevencionista","Jefe de obra","Jefe administrativo"],"equipo")

      set_ponderation_values(output[0])
      output.write('output.xlsx')
    # rescue Exception => e

    #   byebug
    end

  end
end
