#!/usr/bin/env nu

# Render a deck to images for visual review.
#
#   render.nu <deck.pptx> [-o out-dir] [-d dpi] [-c columns]
#
# Converts the deck with LibreOffice, rasterises every page with pdftoppm,
# and builds contact sheets with montage. It also prints the fonts the PDF
# actually embeds, which is the only reliable way to spot a silent font
# substitution. Output goes to <deck stem>.render/ next to the deck.

def run-tool [label: string, cmd: closure] {
  let result = (do $cmd | complete)
  if $result.exit_code != 0 {
    error make {msg: $"($label) failed: ($result.stderr | str trim)"}
  }
  $result
}

def main [
  deck: path
  --out-dir (-o): path
  --dpi (-d): int = 60
  --columns (-c): int = 4
] {
  if ((which soffice | is-empty) or (which pdftoppm | is-empty) or (which montage | is-empty)) {
    let extra = if ($out_dir | is-empty) { [] } else { [--out-dir $out_dir] }
    ^nix shell nixpkgs#libreoffice nixpkgs#poppler-utils nixpkgs#imagemagick --command $nu.current-exe $env.CURRENT_FILE $deck ...$extra --dpi $dpi --columns $columns
    return
  }

  if not ($deck | path exists) {
    error make {msg: $"no such deck: ($deck)"}
  }

  let stem = ($deck | path parse | get stem)
  let parent = ($deck | path dirname)
  let parent = if ($parent | is-empty) { "." } else { $parent }
  let dir = if ($out_dir | is-empty) { $parent | path join $"($stem).render" } else { $out_dir }
  mkdir $dir
  let stale = ([(glob ($dir | path join "page-*.png")) (glob ($dir | path join "sheet-*.png"))] | flatten)
  if not ($stale | is-empty) { rm --force ...$stale }

  let pdf = ($dir | path join $"($stem).pdf")
  run-tool "libreoffice" { ^soffice --headless --convert-to pdf --outdir $dir $deck }
  if not ($pdf | path exists) {
    error make {msg: $"LibreOffice did not produce ($pdf)"}
  }

  let prefix = ($dir | path join "page")
  run-tool "pdftoppm" { ^pdftoppm -r $dpi -png $pdf $prefix }

  let images = (glob ($dir | path join "page-*.png") | sort)
  if ($images | is-empty) {
    error make {msg: "no page images produced"}
  }

  let per_sheet = $columns * 3
  let sheet_count = ((($images | length) + $per_sheet - 1) / $per_sheet | math floor | into int)
  let sheets = (1..$sheet_count | each {|i|
    let chunk = ($images | skip (($i - 1) * $per_sheet) | take $per_sheet)
    let sheet = ($dir | path join $"sheet-($i).png")
    run-tool "montage" { ^montage ...$chunk -tile $"($columns)x" -geometry +6+6 -background "#DDDDDD" $sheet }
    $sheet
  })

  print $"pages: ($images | length)"
  print $"sheets: ($sheets | str join ', ')"
  print "fonts embedded in the PDF:"
  let fonts = (^pdffonts $pdf | lines | skip 2 | each {|line|
    $line | split row -r '\s+' | first | str trim
  } | uniq)
  for font in $fonts { print $"  ($font)" }
}
