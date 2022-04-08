#! /bin/bash

IFS=; while read -r image; do
    eval "docker pull $image"
done <<< $(cat ./images.txt)
