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
    topics: string[],
    activeTopic: string
}

export default class Topics extends React.Component<TopicsProps, TopicsState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            topics: [],
            activeTopic: ""
        };
    }
    render() {
        if (this.props.title == null || this.state.topics == null)
            return(<div/>);
        return (
            <div className="container">
                <ul className="nav nav-tabs">
                {this.state.topics.map((t) => 
                    <li className="nav-item ">
                        <a data-id= {t} className={"nav-link" + t===this.state.activeTopic ? " active": ""} href="#" onClick={(e) => this.click(e)}>
                            {t}
                        </a>
                    </li>)}
                </ul>
                <TopicFeed topic = {this.state.activeTopic} />
          </div>
                /*
                <Tabs defaultActiveKey={this.state.topics[0]} onSelect={this.handleSelect} id="trending-topics" >
                    {this.state.topics.map((t) => 
                    <Tab eventKey={t} title={t}>
                    </Tab>
                    )}
                    
                </Tabs>*/
           /* <div>
            <ul class="nav nav-tabs">
            {this.state.topics.map((t) => 
            <li data-id = {t} key={t} className={this.state.activeTopic === t ? "on" : "off" } onClick={(e) => this.click(e)}>{t}</li>)}
            </ul>
            <TopicFeed topic = {this.state.activeTopic} />
        </div>*/
            /*<section className='ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500'>
                <h1 className='ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary'>{title}</h1>
                <span className='ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary'>{content}</span>
            </section>*/
        );
    }

    async componentDidMount()
    {
        if(this.props.content == "" && this.props.title == "")
            return;
        await this.fetchTopics();
    }
    async componentDidUpdate(prevProps)
    {
        if(this.props.content === prevProps.contents || this.props.title === prevProps.title || this.props.title === "")
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
        let topics = [];
        if (result.documents.length > 0)
        {
            result.documents.forEach((document) =>
            {
                document.entities.forEach(entity =>
                {
                    topics.push(entity.name);
                })
            });
            this.setState({
                topics : Array.from(new Set(topics)),
                activeTopic: topics[0]
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
