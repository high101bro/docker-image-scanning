#! /bin/sh

images=$(docker images -q)

for image_id in $images; do
	image_name=$(docker images | grep "$image_id" | tr -s ' ' | cut -d ' ' -f 1 | rev | cut -d '/' -f 1 | rev)
	echo '===================================================================================================='
	echo "Scanning: $image_name"
        echo '===================================================================================================='
	eval "grype $image_id --file /home/azureuser/grype-reports/$image_name.txt"
done

