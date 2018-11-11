#!/usr/bin/env nix-shell
#!nix-shell -i bash -p bash wmctrl
trap "exit" INT
CHROME_DIR="$(mktemp -d)"
trap "{ rm -rf $CHROME_DIR; }" EXIT
touch "$CHROME_DIR/First Run"
i=0
readarray lines
echo "log in now"
google-chrome-stable --user-data-dir="$CHROME_DIR" "https://tcgplayer.com" >&/dev/null
for line in "${lines[@]}"; do
  echo "$line"
  google-chrome-stable --user-data-dir="$CHROME_DIR" "$line" >&/dev/null
  i="$((i+1))"
done
