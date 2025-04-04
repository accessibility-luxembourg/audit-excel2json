node_bin="node"
if [[ "$OSTYPE" == "cygwin" ]]; then
        node_bin="node.exe"
elif [[ "$OSTYPE" == "msys" ]]; then
        node_bin="node.exe"
elif [[ "$OSTYPE" == "win32" ]]; then
        node_bin="node.exe"
fi

mkdir -p "$1/json/"
find "$1/" -type f -name "*.xlsx"  | while read i; do echo $i; $node_bin audit_convert.js $i > "$1/json/"`basename $i`.json; done
