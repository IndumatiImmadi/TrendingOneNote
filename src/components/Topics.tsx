import * as React from 'react';
import TopicFeed from './TopicFeed';
//import {Tab, Tabs} from 'react-bootstrap';
//import { keyframes } from '@uifabric/styling';
//import { timingSafeEqual } from 'crypto';

export interface TopicsProps {
    content: string,
    title: string
}

export interface TopicsState {
    topics: {},
    activeTopic: string
}

export default class Topics extends React.Component<TopicsProps, TopicsState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            topics: {},
            activeTopic: ""
        };
    }
    render() {
        if ( !(this.props.title  && this.state.topics) || Object.keys(this.state.topics).length === 0)
            return(<div/>);
        return (
            <div className="container">
                <ul className="nav nav-tabs">
                {Object.keys(this.state.topics).map((t) => 
                    <li className= {"nav-item" + (t===this.state.activeTopic ? " active": "")}>
                        <a data-id= {t} className={"nav-link" + (t===this.state.activeTopic ? " active": "")} href="#" onClick={(e) => this.click(e)}>
                            {t}
                        </a>
                    </li>)}
                </ul>
                <TopicFeed topic = {this.state.activeTopic} wikiUrl = {this.state.topics[this.state.activeTopic].url} />
          </div>
        );
        //style="background-color:rgb(119, 25, 170)" - active
    }

    async componentDidMount()
    {
        if(this.props.content == "" && this.props.title == "")
            return;
        await this.fetchTopics();
    }
    async componentDidUpdate(prevProps)
    {
        if((this.props.content === prevProps.content && this.props.title === prevProps.title) || this.props.title === "")
            return;
        await this.fetchTopics();
    }

    async fetchTopics() {
        let uri = 'https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/entities';
        let accessKey = '46539f915c9e40bda98a854dfeaf8030';
        let response: Response = await fetch(uri, {
        method: "post",
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'Ocp-Apim-Subscription-Key' : accessKey,
        },

        //make sure to serialize your JSON body
        body: JSON.stringify( {'documents' :[
            {  'id' : 2,  'language': 'en', 'text': this.props.title },
            {  'id' : 1,  'language': 'en', 'text': this.props.content }
          ] })
        });
        let result = await response.json();
        let topics = {};
        if (result.documents && result.documents.length > 0)
        {
            result.documents.forEach((document) =>
            {
                document.entities.forEach(entity =>
                {
                    topics[entity.name] = {id: entity.wikipediaId, url: entity.wikipediaUrl};
                })
            });
            this.setState({
                topics : Object.assign({}, topics),
                activeTopic: Object.keys(topics)[0]
            });
        }
    }

    click = (event) =>
    {
        this.setState({
            activeTopic: event.currentTarget.dataset.id
        });
    }
}
