# SharePoint Team Footer Web Part

## Summary

A modern SharePoint Framework (SPFx) web part that displays team organizational structure with interactive features. The web part shows Center Director information, Center Managers with clickable tiles, and detailed team breakdowns when a manager is selected.

![Team Footer Web Part](docs/team-footer-preview.png)

## Features

- **Center Director Section**: Displays static information about the center director from web part properties
- **Center Managers Section**: Shows unique center managers from SharePoint list with:
  - Clickable tiles with hover effects (orange highlight)
  - 3-column responsive grid layout
  - Orange background when selected
- **Team Breakdown Section**: Displays detailed team information when a manager is selected:
  - Team name and description (two-column layout with orange left border)
  - Locations with badge-style display
  - Team Leaders and Tech Leaders with profile pictures

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.20.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Prerequisites

- Node.js v18.17.1 or higher (but less than v19.0.0)
- SharePoint list with the following columns:
  - **Center Manager** (Person or Group) - Single selection
  - **Team Name** (Single line of text)
  - **Team Description** (Multiple lines of text)
  - **Locations** (Choice - Multiple selection)
  - **Team Leaders** (Person or Group - Multiple selection)
  - **Tech Leaders** (Person or Group - Multiple selection)

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| sp-team-footer | Przemek Bok |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | May 26, 2025     | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

1. Clone this repository
2. Navigate to the solution folder
3. In the command-line run:
   ```bash
   npm install
   ```
4. Run the following to build the solution:
   ```bash
   gulp build
   gulp bundle --ship
   gulp package-solution --ship
   ```
5. Deploy the `.sppkg` file from `sharepoint/solution` to your tenant App Catalog
6. Add the web part to a SharePoint page

## Configuration

1. **Create a SharePoint List** with the required columns mentioned in Prerequisites
2. **Add the web part** to your page
3. **Configure the web part**:
   - Select your list from the dropdown
   - Choose the Center Director using the people picker
   - Use the provided links to create a new list or view the selected list

## Development

To run the web part in development mode:

```bash
gulp serve
```

This will open your default browser with the SharePoint workbench where you can test the web part.

## Features in Detail

### Center Director Section
- Displays the selected person's photo, name, job title, and email
- Static tile with no interaction
- Information comes from web part properties

### Center Managers Section
- Retrieves unique center managers from the selected list
- Displays in a responsive 3-column grid (adapts to 2 columns on tablets, 1 on mobile)
- Hover effect: Orange border and top accent
- Click effect: Orange background with white text
- Shows manager's photo, name, job title, and department

### Team Breakdown Section
- Appears only when a center manager is selected
- Shows all teams associated with the selected manager
- Each team displays:
  - **Team Name**: Prominent header
  - **Team Description**: Two-column layout with orange accent lines
  - **Locations**: Badge-style chips for each location
  - **Team Leaders**: Grid of personas with photos and names
  - **Tech Leaders**: Grid of personas with photos and names
- Empty sections are automatically hidden

## Styling

The web part uses:
- Fluent UI React components for consistent Microsoft 365 styling
- Custom SCSS modules for specific design requirements
- Theme-aware colors that adapt to SharePoint themes
- Responsive design for mobile, tablet, and desktop views

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp)
