export class  UserObjectType  {
  public UserID: string;
  public Name: string;
  public ImageUrl:string;
  public DelveImageUrl:string;
  public Department:string;
  public JobTitle:string;
  public Email:string;
  public office:string;
  public UserName:string;
  public ProfileUrl:ProfileUrl;
  public FullName:string;
  public Workphone:string;
}

export class ProfileUrl {
  public LinkUrl:string;
  public DetailsUrl:string;
}