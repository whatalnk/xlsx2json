# xlsx2json
## *.xlsx => *.json (file)
```
xlsx2json -input=data.xlsx -output=data.xlsx.json
```
## *.xlsx => *.json (stdout)
```
xlsx2json -input=data.xlsx -output=-
```
## *.json => *.xlsx
```
xlsx2json -input=data.xlsx.json -output=data.xlsx.json.xlsx
```

## Diffing
*  `.gitattributes`
```
*.xlsx diff=xlsx2json
```
* `.gitconfig` or `.git/config`
```
[diff "xlsx2json"]
  binary = true
  textconv = xlsx2json -input=$1 -output=-
```