{ pkgs ? import <nixpkgs> {} }:

pkgs.mkShell {
  buildInputs = [
    pkgs.python3
    pkgs.python3Packages.pytest
    pkgs.python3Packages.pytest-playwright
    pkgs.python3Packages.openpyxl
    pkgs.python3Packages.requests
    pkgs.python3Packages.python-dotenv
    pkgs.playwright-driver
  ];

  shellHook = ''
    # Set the path to the playwright driver provided by Nix
    export PLAYWRIGHT_BROWSERS_PATH=${pkgs.playwright-driver.browsers}
    export PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1
  '';
}
