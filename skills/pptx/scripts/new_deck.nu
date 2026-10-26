#!/usr/bin/env nu

# Scaffold a deck project from the skill template.
#
#   new_deck.nu <target-dir> [--refresh]
#
# Copies the template (build_deck.py, build.sh, assets/) and the deck library
# into a new project. With --refresh only deck.py is replaced, which is how an
# existing project picks up fixes in the library.

def main [
  target: path
  --refresh
] {
  let here = ($env.CURRENT_FILE | path dirname)
  let template = ($here | path join "../template")
  let library = ($here | path join "deck.py")

  if $refresh {
    cp $library ($target | path join "deck.py")
    print $"refreshed ($target)/deck.py"
    return
  }

  if ($target | path exists) and (ls $target | length) > 0 {
    error make {msg: $"($target) already exists and is not empty"}
  }

  mkdir $target ($target | path join "assets")
  cp ($template | path join "build_deck.py") $target
  cp ($template | path join "build.sh") $target
  cp $library $target
  chmod +x ($target | path join "build.sh")

  print $"scaffolded ($target)"
  print "  build_deck.py   slide functions, edit freely"
  print "  deck.py         reusable library"
  print "  build.sh        python-pptx + pillow via nix shell"
  print "  assets/         logos and images"
}
