#! /bin/sh

images=$(docker images -q)

for image_id in $images; do
	image_name=$(docker images | grep "$image_id" | tr -s ' ' | cut -d ' ' -f 1 | rev | cut -d '/' -f 1 | rev)
        image_version=$(docker images | grep "$image_id" | tr -s ' ' | cut -d ' ' -f 2)

	echo '===================================================================================================='
	echo "Scanning: $image_name:$image_version"
        echo '===================================================================================================='
	eval "grype $image_id --file /home/azureuser/grype/$image_name-$image_version.json --output json --scope Squashed"
done

