import win32com.client

def search_files(query):
    # Create a Search object
    search = win32com.client.Dispatch('Windows.Search')

    # Create a QueryHelper object
    helper = search.CreateQueryHelper()

    # Create a QueryNode object for the search query
    node = helper.ParseStructuredQuery(query)

    # Set the search options
    options = win32com.client.Dispatch('Windows.Search.QueryOptions')
    options.SetResultLimit(50)
    options.SortOrder = 4  # Sort by date modified

    # Execute the search query
    results = search.Execute(node, options)

    # Loop through the search results
    for i in range(results.Count):
        # Get the search result
        result = results.Item(i)

        # Get the file content
        content = result.GetThumbnail()

        # Print the file content
        print(content)

# Example usage: search for files containing the word "Python"
search_files('System.ItemType:".txt" AND contains("Python")')
