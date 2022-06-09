import {WebPartContext} from '@microsoft/sp-webpart-base';
import { SPHttpClient, ISPHttpClientOptions } from "@microsoft/sp-http";

export const getFollowedSites = async (context: WebPartContext) => {
    const responseUrl = `${context.pageContext.web.absoluteUrl}/_api/social.following/my/Followed(types=4)`;
    
    try{
        const response = await context.spHttpClient.get(responseUrl, SPHttpClient.configurations.v1);
        if (response.ok){
            const responseResults = await response.json();
            const unsortedResults =  responseResults.value.map(item => {
                return {
                    title: item.Name,
                    url: item.ContentUri
                };
            });
            return unsortedResults.sort((a, b) => a.title.localeCompare(b.title));
        }else{
            console.log("Response Error: ", response.statusText);
        }
    }catch(error){
        console.log("Error: ", error);
    }


};

export const unFollowSite = async (context: WebPartContext, siteLink: string) => {
    const responseUrl = `${context.pageContext.web.absoluteUrl}/_api/social.following/stopfollowing(ActorType=2,ContentUri=@v,Id=null)?@v='${siteLink}'`;
                                                         
    let spOptions: ISPHttpClientOptions = {
        headers:{
            "Accept": "application/json;odata=nometadata", 
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": "",
        }
    };

    try{
        const response = await context.spHttpClient.post(responseUrl, SPHttpClient.configurations.v1, spOptions);
        if (response.ok){
            console.log("Document is unfollowed successfully", siteLink);
        }else{
            console.log("Document unfollow error : " + siteLink + " " + response.statusText);
        }
    }catch(error){
        console.log('unFollowDocument Error', error);
    }
};

const followSite = async (context: WebPartContext, siteLink: string) => {
    const responseUrl = `${context.pageContext.web.absoluteUrl}/_api/social.following/follow(ActorType=2,ContentUri=@v,Id=null)?@v='${siteLink}'`;

    let spOptions: ISPHttpClientOptions = {
        headers:{
            "Accept": "application/json;odata=nometadata", 
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": "",
        }
    };

    try{
        const response = await context.spHttpClient.post(responseUrl, SPHttpClient.configurations.v1, spOptions);
        if (response.ok){
            console.log("Document is unfollowed successfully", siteLink);
        }else{
            console.log("Document unfollow error : " + siteLink + " " + response.statusText);
        }
    }catch(error){
        console.log('unFollowDocument Error', error);
    }
};