{ pkgs }: {
  deps = [
    pkgs.libreoffice
    pkgs.python311
    pkgs.python311Packages.pip
  ];
}
