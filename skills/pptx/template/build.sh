#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

nix shell --impure --expr \
  'let pkgs = import <nixpkgs> {}; in pkgs.python3.withPackages (ps: [ ps.python-pptx ps.pillow ])' \
  -c python3 build_deck.py
