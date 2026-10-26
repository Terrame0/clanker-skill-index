{
  pkgs,
  inputs,
  ...
}:
with pkgs; [
  nushell
  pandoc
  plantuml
  graphviz
  resvg
  python3
  inputs.idef0-svg-gost.packages.${pkgs.stdenv.hostPlatform.system}.default
]
