import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
//import Header from './Header';
import { HeroListItem } from './HeroList';
import Progress from './Progress';
import Topics from './Topics';

import * as OfficeHelpers from '@microsoft/office-js-helpers';

export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

export interface AppState {
    listItems: HeroListItem[];
    title: string;
    content: string;
}

export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            listItems: [],
            title: "",
            content: ""
        };
    }

    componentDidMount() {
        this.setState({
            listItems: [
                {
                    icon: 'Ribbon',
                    primaryText: 'Achieve more with Office integration'
                },
                {
                    icon: 'Unlock',
                    primaryText: 'Unlock features and functionality'
                },
                {
                    icon: 'Design',
                    primaryText: 'Create and visualize like a pro'
                }
            ]
        });
    }

    render() {
        const {
            title,
            isOfficeInitialized,
        } = this.props;

        if (!isOfficeInitialized) {
            return (
                <Progress
                    title={title}
                    logo='assets/shutterstock.jpg'
                    message='Please sideload your addin to see app body.'
                />
            );
        }

        return (
            <div>
                <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
                <Topics content= {this.state.content} title={this.state.title}/>
            </div>
        );
    }

    click = async () => {
        try {
            await OneNote.run(async context => {
                var page = context.application.getActivePage();
                page.load('title'); 
                var pageContents = page.contents;
                pageContents.load("id,type");

                var outlinePageContents = [];
                var paragraphs = [];
                var richTextParagraphs = [];
                var text = "";
                return context.sync().then(() => 
                {
                    pageContents.items.forEach((pageContent) => 
                    {
                        if(pageContent.type == 'Outline')
                        {
                            pageContent.load('outline,outline/paragraphs,outline/paragraphs/type');
                            outlinePageContents.push(pageContent);
                        }
                    });
                    return context.sync();
                }).then(() =>
                {
                    outlinePageContents.forEach((outlinePageContent) => 
                    {
                        var outline = outlinePageContent.outline;
                        paragraphs = paragraphs.concat(outline.paragraphs.items);
                    });
                    paragraphs.forEach((paragraph) => 
                    {
                        if (paragraph.type == 'RichText') 
                        {
                            richTextParagraphs.push(paragraph);
                            paragraph.load("id,richText/text");
                        }
                    });
                    return context.sync();
                }).then(() =>
                {
                    richTextParagraphs.forEach((richTextParagraph) =>
                    {
                        text += "\n" + richTextParagraph.richText.text;
                    })
                    this.setState({
                        content : text,
                        title: page.title
                    });
                }).catch((error) => 
                {
                    console.log("Error: " + error); 
                    if (error instanceof OfficeExtension.Error) { 
                        console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                    } 
                });
            });
        } catch(error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        };
    }
}
