defmodule Elixlsx.Native do
  @app Mix.Project.config()[:app]
  @compile {:autoload, false}
  @on_load :load_nif
  def load_nif() do
    file =
      :code.lib_dir(@app) ++
        '/priv/precompiled/' ++ :erlang.system_info(:system_architecture) ++ '/libelixlsx_native'

    :erlang.load_nif(file, 0)
  end

  def write_excel_nif(_workbook) do
    :erlang.nif_error("nif not loaded")
  end

  def write_excel(workbook) do
    files = write_excel_nif(workbook)
    :zip.create('workbook.xlsx', files, [:memory])
  end
end
