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

  def write_excel(_workbook) do
    :erlang.nif_error("nif not loaded")
  end
end

fn ->
  alias Elixlsx.{Workbook, Sheet}

  rows =
    Enum.map(1..300_0, fn _ ->
      Enum.map(1..10, fn _ -> Base.encode16(:crypto.strong_rand_bytes(5)) end)
    end)

  workbook = %Workbook{sheets: [%Sheet{name: "sheet-1", rows: rows}]}
  :timer.tc(fn -> Elixlsx.Native.write_excel(workbook) end)
  :timer.tc(fn -> Elixlsx.write_to_memory(workbook, "filename.xlsx") end)
end
