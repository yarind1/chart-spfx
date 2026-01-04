import * as React from "react";
import { IListItem } from "../../../services/SharePoint/IListItem";
import SharePointSerivce  from "../../../services/SharePoint/SharePointService";


export interface IChartProps {
    chartTitle: string;

}
export interface IChartState {
    items: IListItem[];
    loading: boolean;
    error: string | null;
}

export default class Chart extends React.Component<IChartProps, IChartState> {
    constructor(props: IChartProps) {
        super(props);  

        //bind methods
        this.getItems = this.getItems.bind(this);

        //set initial state
        this.state = {
            items: [],
            loading: false,
            error: null,
        };
    }   


    public render():JSX.Element {
        return (
            <div>
                <h1>{this.props.chartTitle}</h1>
                {this.state.error && <p>Error: {this.state.error}</p>}
                <ul>
                    {this.state.items.map((item)=>{
                        return <li key={item.Id}>{item.Title}</li>;
                    })}
                </ul>

                <button onClick={this.getItems} disabled={this.state.loading}>{this.state.loading ? 'Loading...' : 'Refresh'}</button>
            </div>
        );
    }

    public getItems():void {
        this.setState({loading:true});
        SharePointSerivce.getListItems("54df84c3-1aeb-4ce4-9ba4-4f854bef6fda").then(
            (items)=>{
            this.setState({
                items: items.value,
                loading:false,
                error:null,
            });
        }).catch((error)=>{
            this.setState({
                error: error.message,
                loading:false,
            });
        });
    }
}

