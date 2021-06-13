for file in talend_*20170410.*
do
    mv -i "${file}" "${file/-m20170410-/-m20170411-}"
done