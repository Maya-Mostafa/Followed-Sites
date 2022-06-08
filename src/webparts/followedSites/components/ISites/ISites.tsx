import * as React from "react";
import styles from '../FollowedSites.module.scss';
import { ISitesProps } from "./ISitesProps";
import { TooltipHost, IconButton } from 'office-ui-fabric-react';

export const ISites = (props: ISitesProps) => {

    const filteredDocItems = 
                            props.searchTxt ? 
                                (props.documentItems ? props.documentItems.filter(item => item.title.toLowerCase().indexOf(props.searchTxt) >= 0) : []) 
                            : props.documentItems;

	return (
        <div className={styles.linkCntnr}>
            {filteredDocItems.map((docItem: any) => {
                return(
                    <>
                        {props.editEnabled ?
                            <div className={styles.toggleNdTxt} key={docItem.id}>
                                <TooltipHost
                                    content="Delete"
                                    id={`tooltip-${docItem.id}`}
                                    calloutProps={{ gapSpace: 0 }}
                                    setAriaDescribedBy={false}
                                >
                                    <IconButton className={styles.deleteIcon} iconProps={{iconName: 'Delete'}} aria-label="Delete" onClick={ () => props.unFollowHandler(docItem.url)}/>
                                </TooltipHost>
                                <a 
                                    className={styles.linkChk}
                                    key={docItem.id} 
                                    href={docItem.url}>
                                        {docItem.title}
                                </a>
                            </div>
                            :
                            <div className={styles.toggleNdTxt} key={docItem.id}>
                                <a 
                                    className={styles.linkChk}
                                    key={docItem.id} 
                                    href={docItem.url}>
                                        {docItem.title}
                                </a>
                            </div>
                        }
                    </>
                    
                );
            })}
        </div>
	);
};

