import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'

class ErrorBoundary extends React.Component {
  constructor(props){super(props);this.state={hasError:false,error:null};}
  static getDerivedStateFromError(e){return{hasError:true,error:e};}
  render(){
    if(this.state.hasError)return <div style={{padding:40,textAlign:"center",color:"red"}}>
      <h2>Hata oluştu</h2><pre style={{fontSize:12,textAlign:"left"}}>{this.state.error?.message}</pre>
    </div>;
    return this.props.children;
  }
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <ErrorBoundary><App /></ErrorBoundary>
)
