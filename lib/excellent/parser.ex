defmodule Excellent.Parser do
  require Record

  Record.defrecord :xmlElement, Record.extract(:xmlElement, from_lib: "xmerl/include/xmerl.hrl")
  Record.defrecord :xmlAttribute, Record.extract(:xmlAttribute, from_lib: "xmerl/include/xmerl.hrl")
  Record.defrecord :xmlText, Record.extract(:xmlText, from_lib: "xmerl/include/xmerl.hrl")

  @shared_string_type 's'
  @number_type 'n'
  @boolean_type 'b'

  @true_value '1'
  @false_value '0'

  def sax_parse_worksheet(xml_content, shared_strings, styles) do
    {:ok, res, _} = :xmerl_sax_parser.stream(
      xml_content,
      event_fun: &event/3,
      event_state: %{
        shared_strings: shared_strings,
        styles: styles,
        col: 0,
        row: 0,
        prev_col: 0,
        prev_row: 0,
        content: []
      }
    )
    Enum.reverse(res[:content])
  end

  def parse_worksheet_names(xml_content) do
    {xml, _rest} = :xmerl_scan.string(:erlang.binary_to_list(xml_content))
    :xmerl_xpath.string('/workbook/sheets/sheet/@name', xml)
      |> Enum.map(fn(x) -> :erlang.list_to_binary(xmlAttribute(x, :value)) end)
      |> List.to_tuple
  end

  def parse_shared_strings(xml_content) do
    {xml, _rest} = :xmerl_scan.string(:erlang.binary_to_list(xml_content))
    :xmerl_xpath.string('/sst/si/t', xml)
      |> Enum.map(fn(element) -> xmlElement(element, :content) end)
      |> Enum.map(fn(texts) -> Enum.map(texts, fn(text) -> to_string(xmlText(text, :value)) end) |> Enum.join end)
      |> List.to_tuple
  end

  def parse_styles(xml_content) do
    {xml, _rest} = :xmerl_scan.string(:erlang.binary_to_list(xml_content))
    lookup = :xmerl_xpath.string('/styleSheet/numFmts/numFmt', xml)
      |> Enum.map(fn(numFmt) -> { extract_attribute(numFmt, 'numFmtId'), extract_attribute(numFmt, 'formatCode') } end)
      |> Enum.into(%{})

    :xmerl_xpath.string('/styleSheet/cellXfs/xf/@numFmtId', xml)
      |> Enum.map(fn(numFmtId) -> to_string(xmlAttribute(numFmtId, :value)) end)
      |> Enum.map(fn(numFmtId) -> lookup[numFmtId] end)
      |> List.to_tuple
  end

  defp extract_attribute(node, attr_name) do
    [ret | _] = :xmerl_xpath.string('./@#{attr_name}', node)
    xmlAttribute(ret, :value) |> to_string
  end

  defp calculate_type(style, type) do
    stripped_style = Regex.replace(~r/(\"[^\"]*\"|\[[^\]]*\]|[\\_*].)/i, style, "")
    if Regex.match?(~r/[dmyhs]/i, stripped_style) do
      "date"
    else
      case {type, style} do
        {@shared_string_type, _} ->
          "shared_string"
        {@number_type, _} ->
          "number"
        {@boolean_type, _} ->
          "boolean"
        _ ->
          "string"
      end
    end
  end

  defp event({:startElement, _, 'row', _, _}, _, state) do
    Dict.merge(state, %{current_row: [], col: 0, prev_col: 0})
  end

  defp event({:startElement, _, 'c', _, [{_, _, 'r', col_row}, {_, _, @shared_string_type, style},
                                         {_, _, 't', type}]}, _, state) do
    {col, row} = parse_col_row(col_row)
    {style_int, _} = :string.to_integer(style)
    style_content = elem(state.styles, style_int)
    type = calculate_type(style_content, type)
    Dict.merge(state, %{type: type, col: col, row: row, prev_col: state.col, prev_row: state.row})
  end

  defp event({:startElement, _, 'c', _, [{_, _, 'r', col_row}, _, {_, _, 't', 's'}]}, _, state) do
    {col, row} = parse_col_row(col_row)
    Dict.merge(state, %{type: "string", col: col, row: row, prev_col: state.col, prev_row: state.row})
  end

  defp event({:endElement, _, 'c', _}, _, state) do
    Dict.delete(state, :type)
  end

  defp event({:startElement, _, 'v', _, _}, _, state) do
    Dict.put(state, :collect, true)
  end

  defp event({:endElement, _, 'v', _}, _, state) do
    Dict.put(state, :collect, false)
  end

  defp event({:characters, chars}, _, %{ collect: true, type: type } = state) do
    value = to_string(chars) |> Type.from_string(%{type: type, lookup: state[:shared_strings]})
    %{state | current_row: filler(state.prev_col, state.col) ++ [value | state[:current_row]]}
  end

  defp event({:endElement, _, 'row', _}, _, state) do
    %{state | content: [Enum.reverse(state[:current_row]) | state[:content]], prev_col: 0}
  end

  defp event(_, _, state),
    do: state

  defp filler(prev_col, col) do
    Stream.repeatedly(fn -> nil end) |> Enum.take(col - prev_col - 1)
  end

  defp parse_col_row(col_row),
    do: parse_col_row([], [], col_row)
  defp parse_col_row(col, row, []),
    do: {dec_b26(col), :string.to_integer(row) |> elem(0)}
  defp parse_col_row(col, row, [h|t]) when h in ?A..?Z,
    do: parse_col_row(col ++ [h], row, t)
  defp parse_col_row(col, row, [h|t]) when h in ?0..?9,
    do: parse_col_row(col, row ++ [h], t)

  defp dec_b26(digits) when is_list(digits), do: dec_b26(Enum.reverse(digits), 1, 0)
  defp dec_b26([d | rest], m, acc), do: dec_b26(rest, m * 26, dec_b26_digit(d) * m + acc)
  defp dec_b26([], _, acc), do: acc
  defp dec_b26_digit(digit), do: digit - ?A + 1
end
