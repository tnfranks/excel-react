import React from 'react';
import ReactDOM from 'react-dom';
import App from './App';
import registerServiceWorker from './registerServiceWorker';
import styled from 'styled-components';

const Body = styled.body`
    width: 95vw;
    margin: auto;
    padding: 0;
    font-family: sans-serif;
`

const Office = window.Office;

Office.initialize = () => {
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        ReactDOM.render(
            <p>
                Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.
            </p>,
            document.getElementById('root')
        )
    }
    ReactDOM.render(<Body><App /></Body>, document.getElementById('root'))
};
registerServiceWorker();
