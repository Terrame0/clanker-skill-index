#!/usr/bin/env nu

const TOC_TITLE = "СОДЕРЖАНИЕ"

# Build a GOST-formatted .docx from a markdown report.
#
#   build.nu <report.md> [out.docx] [--titlepage <titlepage>] [--toc]
#
# The optional title page is either a markdown fragment (custom-style divs,
# concatenated in front of the report by pandoc) or a .docx (its paragraphs are
# spliced in verbatim after pandoc, so the layout is reproduced exactly).
# `--toc` inserts a СОДЕРЖАНИЕ page with a native Word TOC field, which the
# reader's editor fills in on open. Diagrams referenced from the markdown
# (e.g. diagrams/org.png) must be rendered first with render_diagrams.nu.
# Missing tooling is pulled in through `nix shell` automatically.

def main [
  src: path
  out?: path
  --titlepage (-t): path
  --toc (-T)
] {
  if ((which pandoc | is-empty) or (which python3 | is-empty)) {
    let pos = if $out == null { [$src] } else { [$src $out] }
    let tp = if $titlepage == null { [] } else { [--titlepage $titlepage] }
    let toc = if $toc { [--toc] } else { [] }
    ^nix shell nixpkgs#pandoc nixpkgs#python3 --command $nu.current-exe $env.CURRENT_FILE ...$pos ...$tp ...$toc
    return
  }

  let script_dir = ($env.CURRENT_FILE | path dirname)
  let final = if $out == null {
    ($src | str replace --regex '\.md$' '') + ".docx"
  } else {
    $out
  }

  let is_docx = ($titlepage != null) and ($titlepage | str ends-with ".docx")
  let inputs = if ($titlepage != null) and (not $is_docx) {
    [$titlepage $src]
  } else {
    [$src]
  }

  let work = (^mktemp -d | str trim)
  let base_docx = ($work | path join "base.docx")
  let ref_docx = ($work | path join "reference_gost.docx")
  let raw_docx = ($work | path join "out.docx")
  let reference_py = ($script_dir | path join "gost_reference.py")
  let postprocess_py = ($script_dir | path join "gost_postprocess.py")

  let toc_args = if $toc {
    [--toc --toc-depth=3 --metadata $"toc-title=($TOC_TITLE)"]
  } else {
    []
  }
  let post_toc = if $toc { [--toc] } else { [] }

  try {
    ^pandoc --print-default-data-file reference.docx o> $base_docx
    ^python3 -B $reference_py $ref_docx $base_docx
    ^pandoc ...$inputs ...$toc_args --no-highlight --reference-doc $ref_docx -o $raw_docx
    if $is_docx {
      ^python3 -B $postprocess_py $raw_docx $final --titlepage $titlepage ...$post_toc
    } else {
      ^python3 -B $postprocess_py $raw_docx $final ...$post_toc
    }
    print $"written: ($final)"
  } finally {
    rm -rf $work
  }
}
