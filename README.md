# ElixlsxNative

rust impl for elixlsx writer

## Usage

First adding `elixlsx_native` to your list of dependencies in `mix.exs`:

```elixir
def deps do
  [
    {:elixlsx_native, git: "~> 0.1.0"}
  ]
end
```
```
iex(1)>   alias Elixlsx.{Workbook, Sheet}
[Elixlsx.Workbook, Elixlsx.Sheet]
iex(2)>   rows =
...(2)>     Enum.map(1..10000, fn _ ->
...(2)>       Enum.map(1..10, fn _ -> Base.encode16(:crypto.strong_rand_bytes(5)) end)
...(2)>     end)
[
  ...
]
iex(3)>   workbook = %Workbook{sheets: [%Sheet{name: "sheet-1", rows: rows}]}
%Elixlsx.Workbook{
  datetime: nil,
  sheets: [
    %Elixlsx.Sheet{
      col_widths: %{},
      merge_cells: [],
      name: "sheet-1",
      pane_freeze: nil,
      row_heights: %{},
      rows: [
        ...
      ],
      show_grid_lines: true
    }
  ]
}
iex(4)>   :timer.tc(fn -> Elixlsx.Native.write_excel(workbook) end)
{349123,
 {:ok,
  {'workbook.xlsx',
   <<80, 75, 3, 4, 20, 0, 0, 0, 8, 0, 204, 88, 65, 77, 187, 224, 213, 73, 201,
     0, 0, 0, 85, 1, 0, 0, 16, 0, 0, 0, 100, 111, 99, 80, 114, 111, 112, 115,
     47, 97, 112, 112, 46, 120, ...>>}}}
iex(5)>   :timer.tc(fn -> Elixlsx.write_to_memory(workbook, "workbook.xlsx") end)
{2351459,
 {:ok,
  {'workbook.xlsx',
   <<80, 75, 3, 4, 20, 0, 0, 0, 8, 0, 206, 88, 65, 77, 49, 105, 140, 73, 201, 0,
     0, 0, 85, 1, 0, 0, 16, 0, 0, 0, 100, 111, 99, 80, 114, 111, 112, 115, 47,
     97, 112, 112, 46, 120, ...>>}}}
     
```