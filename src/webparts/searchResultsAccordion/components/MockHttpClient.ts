import { QAList } from '../models/IQALists';

export default class MockHttpClient  {

  private static _items: QAList[] = [{ Title: 'Mock List Question 1', bcAnswer: 'Mock List Answer 1',SiteTitle:'Mock Site Title',SiteUrl:'Mock Site Url' },
                                      { Title: 'Mock List Question 2', bcAnswer: 'Mock List Answer 1',SiteTitle:'Mock Site Title',SiteUrl:'Mock Site Url' },
                                      { Title: 'Mock List Question 3', bcAnswer: 'Mock List Answer 1',SiteTitle:'Mock Site Title',SiteUrl:'Mock Site Url' }];

  public static get(): Promise<QAList[]> {
    return new Promise<QAList[]>((resolve) => {
      resolve(MockHttpClient._items);
    });
  }
}