import {WebPartContext} from '@microsoft/sp-webpart-base';
import { SPHttpClient, ISPHttpClientOptions } from "@microsoft/sp-http";

export const getFollowedDocuments = async (context: WebPartContext) => {
    const responseUrl = `${context.pageContext.web.absoluteUrl}/_api/social.following/my/Followed(types=2)`;

    try{
        const response = await context.spHttpClient.get(responseUrl, SPHttpClient.configurations.v1);
        if (response.ok){
            const responseResults = await response.json();
            return responseResults.value.map(item => {
                return {
                    title: item.Name,
                    url: item.ContentUri
                };
            });
        }else{
            console.log("Response Error: ", response.statusText);
        }
    }catch(error){
        console.log("Error: ", error);
    }


};

export const unFollowDocument = async (context: WebPartContext, docLink: string) => {
    const responseUrl = `${context.pageContext.web.absoluteUrl}/_api/social.following/stopfollowing(ActorType=1,ContentUri=@v,Id=null)?@v='${docLink}'`;

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
            console.log("Document is unfollowed successfully", docLink);
        }else{
            console.log("Document unfollow error : " + docLink + " " + response.statusText);
        }
    }catch(error){
        console.log('unFollowDocument Error', error);
    }
};

const followDocument = async (context: WebPartContext, docLink: string) => {
    const responseUrl = `${context.pageContext.web.absoluteUrl}/_api/social.following/follow(ActorType=1,ContentUri=@v,Id=null)?@v='${docLink}'`;

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
            console.log("Document is unfollowed successfully", docLink);
        }else{
            console.log("Document unfollow error : " + docLink + " " + response.statusText);
        }
    }catch(error){
        console.log('unFollowDocument Error', error);
    }
};