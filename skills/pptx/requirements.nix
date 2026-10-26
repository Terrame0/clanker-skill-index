{
  pkgs,
  ...
}:
with pkgs; [
  (python3.withPackages (ps: [
    ps.python-pptx
    ps.pillow
  ]))
  nushell
  libreoffice
  poppler-utils
  imagemagick
  nerd-fonts.jetbrains-mono
  dejavu_fonts
  liberation_ttf
]
