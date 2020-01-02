# NewFeatureAnnouncements
Azure Function to send New Feature Announcements.

The MSTeams Azure DevOps repo has several webhooks installed that notify this function when a feature's ring enablement fields are changed. This function determines if the feature has been enabled in a ring, makes a messageCard with the feature's details, and sends it to all of the Teams incoming webhooks that are subscribed to it.
