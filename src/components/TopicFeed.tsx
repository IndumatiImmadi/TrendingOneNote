import * as React from 'react';
import Progress from './Progress';
import * as Fetch from '../utilities/BingApi'
//import { TokenStorage } from '@microsoft/office-js-helpers';
//import { timingSafeEqual } from 'crypto';

export interface TopicFeedProps {
    topic: string
    wikiUrl: string
}

export interface TopicFeedEntry {
    title: string,
    url: string,
    imageUrl: string,
    description: string
}

export interface TopicFeedState {
    topicNewsFeed: {[url:string]: TopicFeedEntry },
    topicSearchFeed: {[url:string]: TopicFeedEntry},
    activeFeed: string
}

export default class TopicFeed extends React.Component<TopicFeedProps, TopicFeedState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            topicNewsFeed: {},
            topicSearchFeed: {},
            activeFeed: 'topicSearchFeed'
        };
    }

    render() {
        if (this.state.topicSearchFeed == {})
            return(<Progress
                title=""
                logo='assets/logo-filled.png'
                message='Select a topic'
            />);
        else
        {
            return(
              <div className="container">
                <h4>Topic</h4> 
                {this.tabContent(this.props.wikiUrl, this.state.topicSearchFeed[this.props.wikiUrl])}
                <h4>Search Results: {this.props.topic}</h4>
                <ul className="nav nav-tabs">
                    <li className="nav-item ">
                        <a data-id= 'Web' className={"nav-link" + ('topicSearchFeed'===this.state.activeFeed ? " active": "")} href="#" onClick={() => 'topicSearchFeed'!==this.state.activeFeed ? this.setState({activeFeed: 'topicSearchFeed'}): {}}>
                            Web
                        </a>
                    </li>
                    <li className="nav-item ">
                        <a data-id= 'News' className={"nav-link" + ('topicNewsFeed'===this.state.activeFeed ? " active": "")} href="#" onClick={() => 'topicNewsFeed'!==this.state.activeFeed ? this.setState({activeFeed: 'topicNewsFeed'}): {}}>
                            Trending News
                        </a>
                    </li>
                </ul>  
              {Object.keys(this.state[this.state.activeFeed]).map((articleUrl)  => 
                    this.tabContent(articleUrl, this.state[this.state.activeFeed][articleUrl])
                )}
              </div>
            );
        }
    }

    async componentDidMount()
    {
        if (this.props.topic == "" || this.props.topic == null)
        {
            return;
        }
        await this.fetchSearchFeed();
    }
    async componentDidUpdate(prevProps)
    {
        if(this.props.topic == "")
            return;
        if (this.props.topic !== prevProps.topic)
        {
            this.state = {
                topicNewsFeed: {},
                topicSearchFeed: {},
                activeFeed: 'topicSearchFeed'
            };
        }
        if (this.state.activeFeed === 'topicSearchFeed' && Object.keys(this.state.topicSearchFeed).length === 0)
            await this.fetchSearchFeed();
        if (this.state.activeFeed === 'topicNewsFeed' && Object.keys(this.state.topicNewsFeed).length === 0)
            await this.fetchNewsFeed();
        //await this.fetchTweets()
    }

    async fetchNewsFeed() {
        let newsResults = await Fetch.CallBingSearchApi(this.props.topic, "news")
        this.setState({topicNewsFeed : newsResults});
    }

    async fetchSearchFeed() {
        let newsResults = await Fetch.CallBingSearchApi(this.props.topic)
        this.setState({topicSearchFeed : newsResults});
    }

    clickAddtoNote = (event) =>
    {
        var clickedData = this.state[this.state.activeFeed][event.currentTarget.dataset.id];
        OneNote.run(async context => {
            var page = context.application.getActivePage();
            page.load('title'); 
            var pageContents = page.contents;
            pageContents.load("id,type");
            page.addOutline(200, 200, `<div className="card ">
                <div className="card-body">
                <h4 className="card-title"><a href=${clickedData.url}>${clickedData.title}</a></h4>
                <p className="card-text">${clickedData.description}</p>
                </div>
                <br/>
            </div>`);
            
            return context.sync().catch((error) => 
            {
                console.log("Error: " + error); 
                if (error instanceof OfficeExtension.Error) { 
                    console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                } 
            });
        });
    }

    tabContent = (articleUrl:string, article: TopicFeedEntry) => {
        if (!article)
            return ( <br/> );
        return(<div className="card">
                        {article.imageUrl ? 
                            <img className="card-img-top" src={article.imageUrl} style={{width:'100%'}}/> 
                            : <div/>}                        
                        <div className="card-body">
                        <h4 className="card-title"><a href={articleUrl}>{article.title}</a></h4>
                        <p className="card-text">{article.description}</p>
                        <button className="btn btn-primary" data-id={articleUrl} onClick={(e)=>this.clickAddtoNote(e)}>Insert Link</button>
                        </div>
                        <br/>
                        </div> )
    }



}
