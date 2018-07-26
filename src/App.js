import React, { Component } from 'react'
import styled from 'styled-components'
import { DefaultButton } from 'office-ui-fabric-react/lib/Button'
import { TextField } from 'office-ui-fabric-react/lib/TextField'
import axios from 'axios'


const Header = styled.div`
  width: 100%;
  background: black;
  color: #fff;
  position: absolute;
  top: 0;
  left: 0;
  padding: 15px;
  height: 80px;
  overflow: hidden;
`

const Main = styled.div`
  background: #fff;
  position: fixed;
  top: 80px;
  left: 0;
  right: 0;
  bottom: 0;
  padding: 15px;
  overflow: auto;
`

class App extends Component {

  state = {
    loading: false,
    inputValue: '',
    data: []
  }

  onInputChanged = (value) => {
    this.setState({ inputValue: value })
  }

  onColorMe = () => {
    window.Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = 'mediumseagreen';
      await context.sync();
    });
  }

  getData = (event) => {
    event.preventDefault()
    this.setState({ loading: true, data: [] }, () => {
      return axios.get(`https://moonmen-server.herokuapp.com/search/${this.state.inputValue}?radius=10`)
        .then(resp => this.setState({
          loading: false,
          data: [[].concat(resp.data.data[0].id)]
        }))
    })

    // window.Excel.run(async (context) => {
    //   const range = context.workbook.getSelectedRange()
    //   range.values = [[].concat(inputValue)];
    //   await context.sync();
    // })
  }

  createTable = () => {
    window.Excel.run(context => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet()
      const expensesTable = currentWorksheet.tables.add('A1:A1', true)
      expensesTable.name = 'ExpensesTable'

      expensesTable.getHeaderRowRange().values = [['id']]
      //[["Date", "Merchant", "Category", "Amount"]]

      const tableData = this.state.data
      // [
      //   ["1/1/2017", "The Phone Company", "Communications", "120"],
      //   ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
      //   ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
      //   ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
      //   ["1/11/2017", "Bellows College", "Education", "350.1"],
      //   ["1/15/2017", "Trey Research", "Other", "135"],
      //   ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
      // ]

      expensesTable.rows.add(null, tableData)

      //expensesTable.columns.getItemAt(3).getRange().numberFormat = [['$#,##0.00']];
      expensesTable.getRange().format.autofitColumns();
      expensesTable.getRange().format.autofitRows();

      return context.sync()
    })
      .catch(error => {
        if (error instanceof window.OfficeExtension.Error) {
          console.log(`Debug info: ${JSON.stringify(error.debugInfo)}`)
        }
      })
  }

  render() {
    let loader = <div></div>
    if(this.state.loading) {
      loader = <div>Loading...</div>
    } else {
      loader = <div>{this.state.data}</div>
    }

    return (
      <div id="content">
        <Header>
          <h1>Welcome</h1>
        </Header>
        <Main>
          <form onSubmit={this.getData}>
            <TextField type="text" placeholder="Search..." value={this.state.inputValue} onChanged={this.onInputChanged} />
          </form>
          <br />
          <DefaultButton type="submit" primary onClick={this.createTable}>Create Table</DefaultButton>
          <br />
          <br />
          {loader}
        </Main>
      </div>
    );
  }
}

export default App;
