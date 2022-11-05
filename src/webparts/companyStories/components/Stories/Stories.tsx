import { Persona, PersonaSize, Text } from "office-ui-fabric-react";
import * as React from "react";
import { LivePersona } from "@pnp/spfx-controls-react/lib/LivePersona";
import { IStoriesProps } from "./IStoriesProps";
import { IStory } from "./IStory";
import styles from "./Stories.module.scss";

const Stories: React.FC<IStoriesProps> = ({ stories, timing, context }) => {
    const [showing, setShowing] = React.useState<number>(0);

    const nextStory: () => void = React.useCallback(() => {
        if (showing + 1 < stories.length) {
            setShowing(showing + 1);
        } else {
            setShowing(0);
        }
    }, [showing, stories]);

    const prevStory: () => void = React.useCallback(() => {
        if (showing - 1 < 0) {
            setShowing(stories.length - 1);
        } else {
            setShowing(showing - 1);
        }
    }, [showing, stories]);

    React.useEffect(() => {
        const timer: number = setTimeout(nextStory, timing * 1000);

        return () => {
            clearTimeout(timer);
        }
    });

    return (
        <div className={styles.container}>
            <div className={styles.stories}>
                <div className={styles.progress}>
                    {stories.map((item: IStory, index: number) => (
                        <div
                            className={`${styles.progressBar} ${index < showing && styles.past}`}
                            style={{
                                transition: showing === index ? `all ${timing}s linear` : 'none',
                                backgroundPosition: showing === index ? `left bottom` : `right bottom`
                            }}
                            key={`${item.Image} - ${index}`} />
                    ))}
                </div>
                {stories.length > 0 &&
                    <div className={styles.story}>
                        <div className={styles.storyAuthor}>
                            <LivePersona
                                serviceScope={/* eslint-disable-line @typescript-eslint/no-explicit-any */ (context.serviceScope as any)}
                                upn={stories[showing].Author.Email}
                                template={
                                    <>
                                        <Persona size={PersonaSize.size24}
                                            text={stories[showing].Author.Name}
                                            imageUrl={stories[showing].Author.Photo}
                                            styles={{ primaryText: { color: 'white', fontWeight: 500, '&:hover': { color: 'whitesmoke' } }}}
                                        />
                                    </>
                                }
                            />
                        </div>
                        <img src={stories[showing].Image} style={{ height: '100%', width: '100%' }} />
                        <div className={styles.left} onClick={prevStory} />
                        <div className={styles.right} onClick={nextStory} />
                        {stories[showing].Content &&
                        <Text className={styles.storyContent}>
                            {stories[showing].Content}
                        </Text>
                        }
                    </div>
                }
            </div>
        </div>
    );
}

export default Stories;
