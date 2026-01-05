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
  Legend,
  ArcElement,    
  LineElement,   
  PointElement   
} from 'chart.js';

import { Bar, Doughnut, Line, Pie } from 'react-chartjs-2';
ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend,
  ArcElement,
  LineElement,
  PointElement
);

export interface IChartProps {
      listId: string;
  selectedFields: string[];
  chartType: string;
  chartTitle: string;
  colors: string[];
}
export interface IChartState {
    items: IListItem[];
    loading: boolean;
    error: string | null;

 //   isShowing:boolean;
}

export default class Chart extends React.Component<IChartProps, IChartState> {
    constructor(props: IChartProps) {
        super(props);  

        //bind methods
        this.getItems = this.getItems.bind(this);
        //this.hideItems = this.hideItems.bind(this);
        this.chartData = this.chartData.bind(this);

        //set initial state
        this.state = {
            items: [],
            loading: false,
            error: null,
       //     isShowing:false,
        };
    }   


    public render():JSX.Element {
 
        return (
            <div>
                <h1>{this.props.chartTitle}</h1>
                {this.state.error && <p>Error: {this.state.error}</p>}
                
                {this.props.chartType === 'Bar' && <Bar data={this.chartData()} />}
                {this.props.chartType === 'Line' && <Line data={this.chartData()} />}
                {this.props.chartType === 'Pie' && <Pie data={this.chartData()} />}
                {this.props.chartType === 'Doughnut' && <Doughnut data={this.chartData()} />}


                <button onClick={this.getItems} disabled={this.state.loading}>{this.state.loading ? 'Loading...' : 'Refresh'}</button>
               
            </div>
        );
    }
// ביטול כפתור "הסתרת הפרויקים" מהרדנור
//<button onClick={this.hideItems} disabled={!this.state.isShowing}>Hide Projects</button>
    public hideItems():void{
        this.setState({
       //     isShowing:false,
            items:[],
        });
    }
    public getItems():void {
        this.setState({loading:true});
        SharePointSerivce.getListItems(this.props.listId).then(
            (items)=>{
            this.setState({
                items: items.value,
                loading:false,
          //      isShowing:true,
                error:null,
            });
        }).catch((error)=>{
            this.setState({
                error: error.message,
                loading:false,
            //    isShowing:false,
            });
        });
    }

    public chartData(): any {
    
        const colors = [this.props.colors[0] || 'rgba(75, 192, 192, 0.6)',
                        this.props.colors[1] || 'rgba(153, 102, 255, 0.6)',
                        this.props.colors[2] || 'rgba(255, 159, 64, 0.6)']  ;
        const data: { labels: string[], datasets: any[] } = {
            labels: [],
            datasets: [],
        };

        this.state.items.forEach((item,i) => {
            const dataset = {
                label:'',
                data: [] as any[],
                backgroundColor: colors[i % colors.length],
                borderColor: colors[i % colors.length]
            };

            //Build dataset
            this.props.selectedFields.forEach((field, index) => {
                let value = item[field];
                if(value === undefined && item[`OData_${field}`] !== undefined){
                    value = item[`OData_${field}`];
                }
                // Add labels
                if (i === 0 && index>0) {
                    data.labels.push(field);
                }

                if(index===0){
                    dataset.label = value;
                }
                else{
                    dataset.data.push(value);
                }
            });
        
            data.datasets.push(dataset);
        }); 
        
        return data;
    }

    
}

