#!/bin/bash

file="src/vimbrewriter.vbs"
tmp1="src/.vimbrewriter.part1.vbs"
tmp2="src/.vimbrewriter.part2.vbs"
outfile="src/.vimbrewriter.new.vbs"

# Run awk to print the line number after the first matching End Sub
split_file_line=$(awk '
    /Sub Undo/ && !found { found = 1; next }
    found && /End Sub/   { print NR + 1; exit }
' "$file")

# Check if we got a valid number
if [[ -n $split_file_line && $split_file_line =~ ^[0-9]+$ ]]; then
    # Optional: display that line
    sed -n "${split_file_line}p" "$file"
else
    echo "No matching Sub Undo / End Sub block found." >&2
    exit 1
fi

head -n $split_file_line "$file" >"$tmp1"

tail -n +$split_file_line "$file" >"$tmp2"

# --- format both parts with vbspretty ---
npx vbspretty "$tmp1" --indentChar "    "
npx vbspretty "$tmp2" --indentChar "    "

# --- fix formatter mistakes ---

# --- rebuild file ---
cat "$tmp1" "$tmp2" >"$outfile"

# --- replace original atomically ---
mv -f "$outfile" "$file"

sed -i 's/ErrorHandler :/ErrorHandler:/g' "$file"
sed -i 's/_ /_/g' "$file"

# --- remove lines that consist only of whitespace ---
sed -i 's/^[[:space:]]\+$//' "$file"

# --- cleanup ---
rm -f "$tmp1" "$tmp2"
