import * as React from 'react';
import styles from './CompanyStories.module.scss';
import { ICompanyStoriesProps } from './ICompanyStoriesProps';
import Stories from './Stories/Stories';
import { IStory } from './Stories/IStory';
import { Spinner, SpinnerSize, Text } from 'office-ui-fabric-react';
import { IStoryReturn } from './IStoryReturn';

const CompanyStories: React.FC<ICompanyStoriesProps> = ({
  title,
  timing,
  sp,
  graph,
  selectedList,
  context
}) => {
  const [loading, setLoading] = React.useState<boolean>(true);
  const [stories, setStories] = React.useState<IStory[]>([]);

  React.useEffect(() => {
    async function get(): Promise<void> {
      setLoading(true);
      const items: IStory[] = await sp.web.lists
        .getById(selectedList)
        .items
        .select("Id", "Image", "Content", "Author0/Id", "Author0/EMail", "Author0/Title")
        .expand("Author0")
        .filter("Active eq 1")
        ().then(async (res: IStoryReturn[]) => {
          const ret: Promise<IStory[]> = Promise.all(res.map(async (item: IStoryReturn): Promise<IStory> => {
            const image: Partial<{serverUrl: string, serverRelativeUrl: string}> = JSON.parse(item.Image);
            const userPhoto: Blob = await graph.users.getById(item.Author0.EMail).photo.getBlob();
            const url: URL | typeof webkitURL = window.URL || window.webkitURL;
            const photoUrl: string = url.createObjectURL(userPhoto);
            
            return {
              Id: item.Id,
              Image: `${image.serverUrl}${image.serverRelativeUrl}`,
              Content: item.Content,
              Author: {
                Id: item.Author0.Id,
                Name: item.Author0.Title,
                Email: item.Author0.EMail,
                Photo: photoUrl
              }
            }
          }));

          return await ret;
        })

      setStories(items);
    }

    Promise.resolve(get().then(() => {
      setLoading(false);
    }).catch(() => setLoading(false))).catch(() => setLoading(false));
  }, [selectedList, graph, sp]);

  return (
    <div className={styles.companyStories}>
      <Text className={styles.title}>
        {title}
      </Text>
      {loading ?
      <Spinner size={SpinnerSize.large} />  
      :
      <Stories stories={stories} context={context} timing={ Number(timing) >= 5 ? Number(timing) : 5 } />
      }
    </div>
  );
}

export default CompanyStories;
