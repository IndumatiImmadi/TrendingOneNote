export async function CallBingSearchApi(searchQuery:string, endpoint?: string) 
{
    let topicFeed = {};
    let uri = `https://api.cognitive.microsoft.com/bing/v7.0/${endpoint ? endpoint + "/" : ""}search?q=${encodeURIComponent(searchQuery)}`;
    let accessKey = 'd55b297774ac4ed2808d1d9774d11923';
    let response: Response = await fetch(uri, 
        {
            method: "get",
            headers: {
                'Ocp-Apim-Subscription-Key': accessKey,
            }
        });
    let result = await response.json();
    result = endpoint ? result : result.webPages
    if (result.value.length > 0) 
    {
        result.value.forEach((article) => 
        {
            var key = article.url || article.thumbnailUrl;
            if (!(key in topicFeed)) 
            {
                topicFeed[key] =
                    {
                        title: article.name,
                        url: key,
                        imageUrl: (article.image && article.image.thumbnail && article.image.thumbnail.contentUrl ? article.image.thumbnail.contentUrl : null) || article.thumbnailUrl,
                        description: article.description ||  article.snippet
                    };
            }
        });
    }
    return topicFeed;
}

export function getWikipediaSnippet(wikiUrl:string, searchResults: {}): string
{
    if(searchResults[wikiUrl])
    {
        return searchResults[wikiUrl].description;
    }
    return null;
}
