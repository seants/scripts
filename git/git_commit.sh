#!/bin/bash

set -e

if [ -a .git/hooks/pre-commit ]
  then
    .git/hooks/pre-commit
fi
args='-m'
read -p "Message: " message
args="${args} \"${message}\""
eval "git commit ${args}"
