import * as React from 'react';
import styles from './SpTeamFooter.module.scss';
import { ISpTeamFooterProps, ICenterManager, ITeamData } from './ISpTeamFooterProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Icon } from '@fluentui/react/lib/Icon';

interface ISpTeamFooterState {
  centerManagers: ICenterManager[];
  selectedManager: ICenterManager | null;
  teamData: ITeamData[];
  loading: boolean;
  error: string;
}

export default class SpTeamFooter extends React.Component<ISpTeamFooterProps, ISpTeamFooterState> {
  constructor(props: ISpTeamFooterProps) {
    super(props);
    this.state = {
      centerManagers: [],
      selectedManager: null,
      teamData: [],
      loading: false,
      error: ''
    };
  }

  public componentDidMount(): void {
    if (this.props.listId) {
      this.loadCenterManagers();
    }
  }

  public componentDidUpdate(prevProps: ISpTeamFooterProps): void {
    if (prevProps.listId !== this.props.listId && this.props.listId) {
      this.loadCenterManagers();
    }
  }

  private async loadCenterManagers(): Promise<void> {
    this.setState({ loading: true, error: '' });

    try {
      const response: SPHttpClientResponse = await this.props.httpClient.get(
        `${this.props.siteUrl}/_api/web/lists(guid'${this.props.listId}')/items?$select=*,CenterManager/Id,CenterManager/Title,CenterManager/EMail,CenterManager/Department,CenterManager/JobTitle&$expand=CenterManager`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        const managersMap = new Map<number, ICenterManager>();
        
        // Get unique center managers
        data.value.forEach((item: any) => {
          if (item.CenterManager && !managersMap.has(item.CenterManager.Id)) {
            managersMap.set(item.CenterManager.Id, {
              id: item.CenterManager.Id,
              title: item.CenterManager.Title,
              email: item.CenterManager.EMail,
              department: item.CenterManager.Department,
              jobTitle: item.CenterManager.JobTitle,
              picture: `/_layouts/15/userphoto.aspx?size=L&accountname=${item.CenterManager.EMail}`
            });
          }
        });

        this.setState({ 
          centerManagers: Array.from(managersMap.values()),
          loading: false 
        });
      }
    } catch (error) {
      console.error('Error loading center managers:', error);
      this.setState({ 
        error: 'Failed to load center managers',
        loading: false 
      });
    }
  }

  private async loadTeamData(manager: ICenterManager): Promise<void> {
    this.setState({ loading: true });

    try {
      const response: SPHttpClientResponse = await this.props.httpClient.get(
        `${this.props.siteUrl}/_api/web/lists(guid'${this.props.listId}')/items?$select=*,CenterManager/Id,TeamLeaders/Id,TeamLeaders/Title,TeamLeaders/EMail,TechLeaders/Id,TechLeaders/Title,TechLeaders/EMail&$expand=CenterManager,TeamLeaders,TechLeaders&$filter=CenterManager/Id eq ${manager.id}`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        const teams: ITeamData[] = data.value.map((item: any) => ({
          id: item.Id,
          teamName: item.TeamName || '',
          teamDescription: item.TeamDescription || '',
          locations: item.Locations || [],
          teamLeaders: item.TeamLeaders || [],
          techLeaders: item.TechLeaders || [],
          centerManager: item.CenterManager
        }));

        this.setState({ 
          teamData: teams,
          loading: false 
        });
      }
    } catch (error) {
      console.error('Error loading team data:', error);
      this.setState({ 
        error: 'Failed to load team data',
        loading: false 
      });
    }
  }

  private handleManagerClick = (manager: ICenterManager): void => {
    this.setState({ selectedManager: manager }, () => {
      this.loadTeamData(manager);
    });
  }

  private renderCenterDirector(): React.ReactElement {
    const { centerDirector } = this.props;
    
    if (!centerDirector) return <></>;

    let directorInfo = null;
    try {
      const parsed = JSON.parse(centerDirector);
      if (parsed && parsed[0]) {
        directorInfo = parsed[0];
      }

      console.log(directorInfo);
    } catch (error) {
      console.error('Error parsing center director:', error);
    }

    if (!directorInfo) return <></>;

    return (
      <div className={styles.section}>
        <h2 className={styles.sectionTitle}>Center Director</h2>
        <div className={styles.directorTile}>
          <Persona
            imageUrl={`/_layouts/15/userphoto.aspx?size=L&accountname=${directorInfo.email}`}
            text={directorInfo.fullName}
            secondaryText={directorInfo.jobTitle || 'Center Director'}
            tertiaryText={''}
            size={PersonaSize.size48}
            imageAlt={directorInfo.text}
          />
        </div>
      </div>
    );
  }

  private renderCenterManagers(): React.ReactElement {
    const { centerManagers, selectedManager } = this.state;

    return (
      <div className={styles.section}>
        <h2 className={styles.sectionTitle}>Center Managers</h2>
        <div className={styles.managerGrid}>
          {centerManagers.map((manager) => (
            <div
              key={manager.id}
              className={`${styles.managerTile} ${selectedManager?.id === manager.id ? styles.selected : ''}`}
              onClick={() => this.handleManagerClick(manager)}
            >
              <Persona
                imageUrl={manager.picture}
                text={manager.title}
                secondaryText={manager.jobTitle}
                tertiaryText={manager.department}
                size={PersonaSize.size48}
                imageAlt={manager.title}
              />
            </div>
          ))}
        </div>
      </div>
    );
  }

  private splitTeamBreakdown(input: string, pageNumber: number): string[]
  {
        // Split the string into an array of lines
        const lines = input.split('\n');
    
        // Calculate the midpoint
        const midpoint = Math.ceil(lines.length / 2);
        
        // Split into two parts
        const firstHalf = lines.slice(0, midpoint);
        const secondHalf = lines.slice(midpoint);
        
        if (pageNumber === 1)
        {
          return firstHalf;
        }
        else if(pageNumber === 2)
        {
          return secondHalf;
        }
        else
        {
          return firstHalf;
        }
  }

  private renderTeamBreakdown(): React.ReactElement | null {
    const { selectedManager, teamData } = this.state;

    if (!selectedManager || teamData.length === 0) return null;

    return (
      <div className={styles.section}>
        <h2 className={styles.sectionTitle}>Team Breakdown</h2>
        {teamData.map((team) => (
          <div key={team.id} className={styles.teamSection}>
            <h3 className={styles.teamName}>{team.teamName}</h3>
            
            {team.teamDescription && (
              <div className={styles.teamDescription}>
                <div className={styles.descriptionColumn}>
                  <div className={styles.columnLine}></div>
                  <div className={styles.columnContent}>
                    {this.splitTeamBreakdown(team.teamDescription,1).map((row) => (
                      <div>
                        {row}<br/>
                      </div>
                    ))}
                  </div>
                </div>
                <div className={styles.descriptionColumn}>
                  <div className={styles.columnLine}></div>
                  <div className={styles.columnContent}>
                  {this.splitTeamBreakdown(team.teamDescription,2).map((row) => (
                      <div>
                        {row}<br/>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}

            {team.locations.length > 0 && (
              <div className={styles.subsection}>
                <h4 className={styles.subsectionTitle}>
                  {/* <Icon iconName="MapPin" className={styles.icon} /> */}
                  Locations
                </h4>
                <div className={styles.locationList}>
                  {team.locations.map((location, index) => (
                    <span key={index} className={styles.location}>{location}</span>
                  ))}
                </div>
              </div>
            )}

            {team.teamLeaders.length > 0 && (
              <div className={styles.subsection}>
                <h4 className={styles.subsectionTitle}>
                  {/* <Icon iconName="People" className={styles.icon} /> */}
                  Team Leaders
                </h4>
                <div className={styles.leaderGrid}>
                  {team.teamLeaders.map((leader) => (
                    <Persona
                      key={leader.Id}
                      imageUrl={`/_layouts/15/userphoto.aspx?size=S&accountname=${leader.EMail}`}
                      text={leader.Title}
                      size={PersonaSize.size32}
                      imageAlt={leader.Title}
                    />
                  ))}
                </div>
              </div>
            )}

            {team.techLeaders.length > 0 && (
              <div className={styles.subsection}>
                <h4 className={styles.subsectionTitle}>
                  {/* <Icon iconName="DeveloperTools" className={styles.icon} /> */}
                  Tech Leaders
                </h4>
                <div className={styles.leaderGrid}>
                  {team.techLeaders.map((leader) => (
                    <Persona
                      key={leader.Id}
                      imageUrl={`/_layouts/15/userphoto.aspx?size=S&accountname=${leader.EMail}`}
                      text={leader.Title}
                      size={PersonaSize.size32}
                      imageAlt={leader.Title}
                    />
                  ))}
                </div>
              </div>
            )}
          </div>
        ))}
      </div>
    );
  }

  public render(): React.ReactElement<ISpTeamFooterProps> {
    const { loading, error } = this.state;
    const { listId } = this.props;

    if (!listId) {
      return (
        <div className={styles.spTeamFooter}>
          <div className={styles.placeholder}>
            <Icon iconName="Info" className={styles.placeholderIcon} />
            <p>Please configure the web part by selecting a list from the property pane.</p>
          </div>
        </div>
      );
    }

    if (error) {
      return (
        <div className={styles.spTeamFooter}>
          <div className={styles.error}>
            <Icon iconName="ErrorBadge" className={styles.errorIcon} />
            <p>{error}</p>
          </div>
        </div>
      );
    }

    return (
      <div className={styles.spTeamFooter}>
        {this.renderCenterDirector()}
        {this.renderCenterManagers()}
        {this.renderTeamBreakdown()}
        {loading && (
          <div className={styles.loading}>
            <Icon iconName="Sync" className={styles.spinner} />
            <span>Loading...</span>
          </div>
        )}
      </div>
    );
  }
}
