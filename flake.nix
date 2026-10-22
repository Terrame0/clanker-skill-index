{
  description = "";
  inputs = {
    sundry-flake.url = "github:Terrame0/sundry";
    nixpkgs.url = "github:NixOS/nixpkgs/nixos-26.05";
    flake-utils.url = "github:numtide/flake-utils";
    idef0-svg-gost = {
      url = "github:Terrame0/IDEF0-SVG-GOST-wrapped";
      inputs.nixpkgs.follows = "nixpkgs";
      inputs.sundry-flake.follows = "sundry-flake";
    };
  };
  outputs = inputs:
    inputs.flake-utils.lib.eachDefaultSystem (system: let
      pkgs = import inputs.nixpkgs {inherit system;};
      sundry = import inputs.sundry-flake {inherit pkgs;};
      inherit (pkgs) lib;
      skills-file-tree = sundry.vfs.dir.from-src ./skills;
    in {
      packages.default = pkgs.symlinkJoin {
        name = "skills-requirements";
        paths = lib.pipe skills-file-tree [
          (sundry.vfs.dir.filter
            (path: file: sundry.vfs.path.get.name path == "requirements.nix"))
          sundry.vfs.dir.load-nix
          (sundry.vfs.dir.collapse
            (path: file: file.expr {inherit pkgs inputs;}))
          lib.concatLists
        ];
      };
      skills = lib.pipe skills-file-tree [
        (sundry.vfs.dir.filter
          (path: file: sundry.vfs.path.get.ext path != "nix"))
        (lib.mapAttrs (dir: file-tree: (sundry.vfs.dir.materialize dir file-tree).drv))
      ];
    });
}
