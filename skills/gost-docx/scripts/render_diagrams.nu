#!/usr/bin/env nu

# Render the diagram sources named on the command line to sibling PNGs.
#
#   render_diagrams.nu <spec>...
#
# A <spec> is a PlantUML path, or an IDEF0 path with an optional level directive
# `#<verb>[=<process>]`:
#
#   foo.puml                      PlantUML             → foo.png            (PLANTUML_DPI, default 300)
#   model.idef0                   schematic (all leaves) → model.svg + .png   (IDEF0_DPI, default 300)
#   model.idef0#focus             focus "<root>"       → model-a0.svg + .png (the A0 context diagram)
#   model.idef0#decompose=<name>  decompose            → model-decompose-<slug>.svg + .png
#   model.idef0#focus=<name>      focus                → model-focus-<slug>.svg + .png
#
# One model file carries every level: render the context with `#focus`, the
# top-level breakdown with `#decompose=<root>`, a level-2 diagram with
# `#decompose=<child>`, and a single process with `#focus=<child>`. Do not split
# a model into one file per level — `toc` prints the process names, and every
# statement is `Noun Verb Noun` in one flat namespace, so a copy per level just
# duplicates statements and drifts.
#
# `<name>` is a process name exactly as `toc` prints it; quote the whole spec
# because it contains `#`, `=` and often spaces. `#` is reserved for the level
# directive and may not appear in a path.
#
# Paths are taken verbatim: the script never scans a directory, so a scratch
# model is rendered — or can break the run — only when it is named explicitly.
#
# The IDEF0 verbs come from IDEF0-SVG-GOST-wrapped (`schematic`, `decompose`,
# `focus`); each produces an SVG that `resvg` rasterizes. The SVG is kept next
# to the model so the PNG can be re-rendered at another DPI without re-running
# the renderer.
#
# Missing tooling is pulled in through `nix shell` automatically; the IDEF0
# binaries come from the wrapper's flake.

def extension [f: string]: nothing -> string {
  $f | path parse | get extension | str downcase
}

# "model.idef0#decompose=Заселение" → "model.idef0"
def idef0-path [spec: string]: nothing -> string {
  $spec | split row '#' | first
}

# "model.idef0#decompose=Заселение" → "decompose"; "model.idef0" → ""
def idef0-verb [spec: string]: nothing -> string {
  let parts = ($spec | split row '#')
  if ($parts | length) <= 1 { return "" }
  $parts | skip 1 | str join '#' | split row '=' | first
}

# "model.idef0#decompose=Заселение" → "Заселение"; else ""
def idef0-process [spec: string]: nothing -> string {
  let full = ($spec | split row '#' | skip 1 | str join '#')
  if not ($full | str contains '=') { return "" }
  $full | split row '=' | skip 1 | str join '='
}

# Collapse runs of whitespace so a process name is a single path component.
def slug [name: string]: nothing -> string {
  $name | str replace --all --regex '\s+' '_'
}

def render-puml [files: list<string>, dpi: string]: nothing -> int {
  if ($files | is-empty) { return 0 }
  ^plantuml -tpng -charset UTF-8 $"-Sdpi=($dpi)" ...$files
  $files | length
}

def render-idef0 [specs: list<string>, dpi: string]: nothing -> int {
  for spec in $specs {
    let model = (idef0-path $spec)
    let verb = (idef0-verb $spec)
    let process = (idef0-process $spec)

    if ($verb not-in ["" "schematic" "decompose" "focus"]) {
      error make {msg: $"render_diagrams.nu: unknown IDEF0 verb '($verb)' in '($spec)'; expected decompose or focus"}
    }

    let dir = ($model | path dirname)
    let stem = ($model | path parse | get stem)
    let has_process = (not ($process | is-empty))

    let name = if ($verb in ["" "schematic"]) {
      $stem
    } else if $verb == "focus" and (not $has_process) {
      $"($stem)-a0"
    } else if $has_process {
      $"($stem)-($verb)-(slug $process)"
    } else {
      $"($stem)-($verb)"
    }
    let svg = ($dir | path join $"($name).svg")

    if ($verb in ["" "schematic"]) {
      open --raw $model | ^schematic o> $svg
    } else if $verb == "decompose" {
      if $has_process {
        open --raw $model | ^decompose $process o> $svg
      } else {
        open --raw $model | ^decompose o> $svg
      }
    } else if $has_process {
      open --raw $model | ^focus $process o> $svg
    } else {
      open --raw $model | ^focus o> $svg
    }

    ^resvg --dpi $dpi $svg ($dir | path join $"($name).png")
  }
  $specs | length
}

def main [...files: string] {
  if ($files | is-empty) {
    error make {msg: "render_diagrams.nu: name at least one .puml or .idef0 file"}
  }

  let unsupported = ($files | where {|f| (extension (idef0-path $f)) not-in ["puml" "idef0"]})
  if not ($unsupported | is-empty) {
    error make {msg: $"render_diagrams.nu: not a .puml or .idef0 path: ($unsupported | str join ', ')"}
  }

  if (which plantuml | is-empty) or (which schematic | is-empty) or (which decompose | is-empty) or (which focus | is-empty) or (which resvg | is-empty) {
    ^nix shell nixpkgs#plantuml nixpkgs#graphviz nixpkgs#resvg github:Terrame0/IDEF0-SVG-GOST-wrapped --command $nu.current-exe $env.CURRENT_FILE ...$files
    return
  }

  let puml_dpi = ($env | get -o PLANTUML_DPI | default "300")
  let idef0_dpi = ($env | get -o IDEF0_DPI | default "300")

  let puml = ($files | where {|f| (extension (idef0-path $f)) == "puml"})
  let idef0 = ($files | where {|f| (extension (idef0-path $f)) == "idef0"})

  let n_puml = (render-puml $puml $puml_dpi)
  let n_idef0 = (render-idef0 $idef0 $idef0_dpi)

  print $"rendered ($n_puml) PlantUML and ($n_idef0) IDEF0 diagram\(s\)"
}
