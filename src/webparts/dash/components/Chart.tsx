import * as React from "react";
import { IListItem } from "../../../services/SharePoint/IListItem";
import SharePointSerivce  from "../../../services/SharePoint/SharePointService";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend
} from 'chart.js';

import { Bar } from 'react-chartjs-2';
ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend
);

export interface IChartProps {
    chartTitle: string;

}
export interface IChartState {
    items: IListItem[];
    loading: boolean;
    error: string | null;
    isShowing:boolean;
}

export default class Chart extends React.Component<IChartProps, IChartState> {
    constructor(props: IChartProps) {
        super(props);  

        //bind methods
        this.getItems = this.getItems.bind(this);
        this.hideItems = this.hideItems.bind(this);

        //set initial state
        this.state = {
            items: [],
            loading: false,
            error: null,
            isShowing:false,
        };
    }   


    public render():JSX.Element {
 
        return (
            <div>
                <h1>{this.props.chartTitle}</h1>
                {this.state.error && <p>Error: {this.state.error}</p>}

                <div>
                    <Bar 
                        data={{
                            labels: ['January', 'February'],
                            datasets: [
                                {
                                    label: 'Apples',
                                    data: [10, 20, 30],
                                    backgroundColor: 'rgba(255, 99, 132, 0.5)', // הוספתי צבע שיהיה ברור
                                },
                                {
                                    label: 'Avocados',
                                    data: [9, 12, 21],
                                    backgroundColor: 'rgba(53, 162, 235, 0.5)',
                                }]
                        }} 
                    />
                </div>

                <ul>
                    {this.state.items.map((item)=>{
                        return <li key={item.Id}>{item.Year}</li>;
                    })}
                </ul>

                <button onClick={this.getItems} disabled={this.state.loading || this.state.isShowing}>{this.state.loading ? 'Loading...' : 'Get Projects'}</button>
                <button onClick={this.hideItems} disabled={!this.state.isShowing}>Hide Projects</button>
            </div>
        );
    }

    public hideItems():void{
        this.setState({
            isShowing:false,
            items:[],
        });
    }
    public getItems():void {
        this.setState({loading:true});
        SharePointSerivce.getListItems("b6a36a73-a877-419e-9653-d2d8bf3a1aa0").then(
            (items)=>{
            this.setState({
                items: items.value,
                loading:false,
                isShowing:true,
                error:null,
            });
        }).catch((error)=>{
            this.setState({
                error: error.message,
                loading:false,
                isShowing:false,
            });
        });
    }
}

