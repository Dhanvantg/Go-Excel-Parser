# Go-Excel-Parser
A GO script that parses <a href="https://docs.google.com/spreadsheets/d/1D0diMLn3Pwwgpcplah0HTppLnEhcSIr0NGgY6Mhgihs/edit?gid=311721139#gid=311721139">this spreadsheet</a> and creates a report customized to the user's flags

### Setup
```go
go build xl.go
./xl --path=gradebook.xlsx --export=json --class=2463
```
