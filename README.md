# xlsx2yaml
Converts xlsx files of a certain format to YAML

If your spreadsheet looks like this:

| name        | style          | rating |
|-------------|----------------|--------|
| hop hog     | ipa            | 8      |
| fanta pants | american amber | 9      |

Then you can xlsx2yaml it:

    xlsx2yaml --xlsx brews.xlsx --sheet Brews --output brews.yaml --beginDataRow 2
    
And then as output you shall receive:

```yaml
Brews:
- style: ipa
  rating: 8
  name: hop hog
- style: american amber
  rating: 9
  name: fanta pants
```
