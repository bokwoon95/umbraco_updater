#!/usr/bin/env bash

# # change to script directory
# cd "${0%/*}"

read -p $'Are you sure you want to delete everything but the latest commit? Enter "YES TAKE ME AWAY" to proceed.\n> ' REPLY
if [ "$REPLY" != "YES TAKE ME AWAY" ]; then
  exit
fi

git checkout --orphan latest_branch
git add -A
git commit -am "Initial Commit"
git branch -D master
git branch -m master
git push -f origin master
echo "History has been rewritten. Initial Commit."
