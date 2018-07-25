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
    topicImageFeed:  {[url:string]: TopicFeedEntry},
    wikiArticle: TopicFeedEntry,
    activeFeed: string
}

export default class TopicFeed extends React.Component<TopicFeedProps, TopicFeedState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            topicNewsFeed: {},
            topicSearchFeed: {},
            topicImageFeed: {},
            wikiArticle: null,
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
                {this.renderWikiArticle()}
                <h4>Search Results: {this.props.topic}</h4>
                <ul className="nav nav-tabs">
                    <li className={"nav-item " + ('topicSearchFeed'===this.state.activeFeed ? " active": "")}>
                        <a data-id= 'Web' className={ "nav-link"  + ('topicSearchFeed'===this.state.activeFeed ? " active": "")} href="#" onClick={() => 'topicSearchFeed'!==this.state.activeFeed ? this.setState({activeFeed: 'topicSearchFeed'}): {}}>
                            Web
                        </a>
                    </li>
                    <li className={"nav-item " + ('topicNewsFeed'===this.state.activeFeed ? " active": "")}>
                        <a data-id= 'News' className={ "nav-link" + ('topicNewsFeed'===this.state.activeFeed ? " active": "")} href="#" onClick={() => 'topicNewsFeed'!==this.state.activeFeed ? this.setState({activeFeed: 'topicNewsFeed'}): {}}>
                            News
                        </a>
                    </li>
                    <li className={"nav-item " + ('topicImageFeed'===this.state.activeFeed ? " active": "")}>
                        <a data-id= 'Images' className={ "nav-link" + ('topicImageFeed'===this.state.activeFeed ? " active": "")} href="#" onClick={() => 'topicImageFeed'!==this.state.activeFeed ? this.setState({activeFeed: 'topicImageFeed'}): {}}>
                            Images
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
            this.setState({
                topicNewsFeed: {},
                topicSearchFeed: {},
                topicImageFeed: {},
                wikiArticle: null,
                activeFeed: 'topicSearchFeed'
            });
        }
        if (this.state.activeFeed === 'topicSearchFeed' && Object.keys(this.state.topicSearchFeed).length === 0)
            await this.fetchSearchFeed();
        if (this.state.activeFeed === 'topicNewsFeed' && Object.keys(this.state.topicNewsFeed).length === 0)
            await this.fetchNewsFeed();
        if (this.state.activeFeed === 'topicImageFeed' && Object.keys(this.state.topicImageFeed).length === 0)
            await this.fetchImageFeed();
        //await this.fetchTweets()
    }

    async fetchNewsFeed() {
        let newsResults = await Fetch.CallBingSearchApi(this.props.topic, "news")
        this.setState({topicNewsFeed : newsResults});
    }

    async fetchSearchFeed() {
        let searchResults = await Fetch.CallBingSearchApi(this.props.topic);
        let wikiArticleResult = searchResults[this.props.wikiUrl];
        delete searchResults[this.props.wikiUrl];
        this.setState({topicSearchFeed : searchResults,
                        wikiArticle: wikiArticleResult });
    }

    async fetchImageFeed() {
        let searchResults = await Fetch.CallBingSearchApi(this.props.topic, "images");
        this.setState({topicImageFeed : searchResults});
    }

    clickAddtoNote = (event) =>
    {
        var clickedData = (this.state.wikiArticle && (this.props.wikiUrl === event.currentTarget.dataset.id)) ?
                            this.state.wikiArticle:
                            this.state[this.state.activeFeed][event.currentTarget.dataset.id];
        OneNote.run(async context => {
            var page = context.application.getActivePage();
            page.load('title'); 
            var pageContents = page.contents;
            pageContents.load("id,type");
            var image = this.state.activeFeed === "topicImageFeed" ?
                     `<img className="card-img-top img-thumbnail" src=${clickedData.imageUrl} style={{width:'100%'}}/>`
                     : "";
            page.addOutline(200, 200, `<div className="card ">
            ${image}
                <div className="card-body">
                <h4 className="card-title"><a href=${clickedData.url}>${clickedData.title}</a></h4>
                <p className="card-text">${clickedData.description ? clickedData.description: ""}</p>
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
                    <p className="card-text"> 
                        <button className="btn btn-info float-right" data-id={articleUrl} onClick={(e)=>this.clickAddtoNote(e)}>
                        <i className="fa fa-plus"></i></button>
                    </p>
                        {article.imageUrl ? 
                            <img className="card-img-top img-thumbnail" src={article.imageUrl} style={{width:'100%'}}/> 
                            : <div/>}                        
                        <div className="card-body">
                        <h4 className="card-title"><a href={articleUrl}>{article.title}</a></h4>
                        <p className="card-text">{article.description}</p>
                        
                        </div>
                        <br/>
                </div> )
    }

    renderWikiArticle()
    {
        if (this.state.wikiArticle && (this.props.wikiUrl === this.state.wikiArticle.url))
            return(
                <div>
                    <h4>Topic</h4> 
                    {this.tabContent(this.props.wikiUrl, this.state.wikiArticle)}
                </div>
            );
        else
            return (<div/>);
    }



}
