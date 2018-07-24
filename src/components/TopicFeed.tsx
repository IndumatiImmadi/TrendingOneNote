import * as React from 'react';
import Progress from './Progress';
//import { TokenStorage } from '@microsoft/office-js-helpers';
//import { timingSafeEqual } from 'crypto';

export interface TopicFeedProps {
    topic: string
}

export interface TopicFeedNewsEntry {
    title: string,
    url: string,
    imageUrl: string,
    description: string
}

export interface TopicFeedState {
    topicNewsFeed: {[url:string]: TopicFeedNewsEntry }
}

export default class TopicFeed extends React.Component<TopicFeedProps, TopicFeedState> {
    token: any;

    constructor(props, context) {
        super(props, context);
        this.state = {
            topicNewsFeed: {}
        };
        /*
        if(OAuth != null)
        {
            
            this.oauth = new OAuth.OAuth2( '09DQco1eZwLA13CmJ7E1UzyFc', 'RBjUllkjgUnTmEGolroCL3lr3IIgdZoA4yNDS95p8w3kyiG7OQ',
            'https://api.twitter.com', null, 'oauth2/token',null
            'https://api.twitter.com/oauth/request_token',
            'https://api.twitter.com/oauth/access_token',
            '09DQco1eZwLA13CmJ7E1UzyFc',
            'RBjUllkjgUnTmEGolroCL3lr3IIgdZoA4yNDS95p8w3kyiG7OQ',
            '1.0A',
            null,
        'HMAC-SHA1');
            this.oauth.getOAuthAccessToken('',{'grant_type': 'client_credentials'}, function(e, access_token){this.token=access_token; console.log(e);})
        }*/
    }

    render() {
        if (this.state.topicNewsFeed == {})
            return(<Progress
                title=""
                logo='assets/logo-filled.png'
                message='Select a topic'
            />);
        else
            return(
              <div className="container">
              {Object.keys(this.state.topicNewsFeed).map((newsUrl)  => 
                <div className="card ">
                        <img className="card-img-top" src={this.state.topicNewsFeed[newsUrl].imageUrl} style={{width:'100%'}}/>
                        <div className="card-body">
                        <h4 className="card-title"><a href={newsUrl}>{this.state.topicNewsFeed[newsUrl].title}</a></h4>
                        <p className="card-text">{this.state.topicNewsFeed[newsUrl].description}</p>
                        <button className="btn btn-primary" data-id={newsUrl} onClick={(e)=>this.clickAddtoNote(e)}>Insert Link</button>
                        </div>
                        <br/>
                </div> )}
              </div>
            );
    }

    async componentDidMount()
    {
        if(!this.token)
        {
            //await this.getTwitterAccessCode();
        }
        if (this.props.topic == "" || this.props.topic == null)
        {
            return;
        }
        await this.fetchFeed();
    }
    async componentDidUpdate(prevProps)
    {
        if(this.props.topic == "" || this.props.topic === prevProps.topic)
            return;
        await this.fetchFeed();
        //await this.fetchTweets()
    }

    async fetchFeed() {
        let uri = 'https://api.cognitive.microsoft.com/bing/v7.0/news/search' + '?q=' + encodeURIComponent(this.props.topic);
        let accessKey = 'd55b297774ac4ed2808d1d9774d11923';
        let response: Response = await fetch(uri, {
        method: "get",
        headers: {
            'Ocp-Apim-Subscription-Key' : accessKey,
        }
        });
        let result = await response.json();
        if (result.value.length > 0)
        {
            var topicNewsFeed = {};
            result.value.forEach((article) =>
            {
                if(!(article.url in topicNewsFeed))
                {
                    topicNewsFeed[article.url] = 
                        {
                            title: article.name,
                            url: article.url,
                            imageUrl: article.image && article.image.thumbnail && article.image.thumbnail.contentUrl ? article.image.thumbnail.contentUrl : null,
                            description: article.description
                        };
                }
            });
            this.setState({topicNewsFeed : topicNewsFeed});
        }
        return result;
    }

    clickAddtoNote = (event) =>
    {
        var clickedData = this.state.topicNewsFeed[event.currentTarget.dataset.id];
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

    async fetchTweets()
    {
        this.token = 'AAAAAAAAAAAAAAAAAAAAAEx47wAAAAAAjTriWLGU3Z2mOYJUOkiqkFG1hLw%3D5zGrks6nSeXaTlshc8lYrxRMaFL54U3YJYx171wRNOcn71iGZb';
        let response = await fetch('https://api.twitter.com/1.1/search/tweets.json?q=nasa&result_type=popular',
        {
            method : "get",
            headers: { Authorization: 'Bearer ' + this.token,
            'Access-Control-Allow-Origin': "*",
            'Access-Control-Allow-Headers':'application/json'},
            mode: 'no-cors'
        });
        let json = await response.json();
        console.log(json);
          /*this.oauth.get(
            'https://api.twitter.com/1.1/search/tweets.json?q=%40twitterapi',
            '128318062-3z1BUbGH8C1F25ISsYmamVRcbRt7THJtvyCsJxBK', //test user token
            '9JNm4WzqphCWBBc6WHNyVW0jurnLf1N4nlU0k98UcTVKv', //test user secret            
            function (e, data, res){
              if (e) console.error(e);        
              console.log(data);
              console.log(res);
            });*/

    }

    async getTwitterAccessCode(){
        let authorization =  btoa(encodeURIComponent("09DQco1eZwLA13CmJ7E1UzyFc") + ":" + encodeURIComponent("RBjUllkjgUnTmEGolroCL3lr3IIgdZoA4yNDS95p8w3kyiG7OQ"))
        let response = await fetch("https://api.twitter.com/oauth2/token", {
            method: "post",
            headers : {
                'Authorization' : `Basic ${authorization}`,
                'Content-Type':'application/x-www-form-urlencoded;charset=UTF-8'
            },
            mode: 'no-cors',
            body: 'grant_type=client_credentials'
            });
        let token = await response.json();
        this.token = token;
    }

}
