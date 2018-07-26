import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';

const MyPage = () => (
  <Fabric>
    <DefaultButton>
      I am a button.
    </DefaultButton>
  </Fabric>
);

export default MyPage;