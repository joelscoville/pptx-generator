from flask import Flask, after_this_request, render_template, request, send_file, jsonify
import pptx_generator_ui
import os

app = Flask(__name__)

@app.route('/')
def home():
   return render_template('index.html')

@app.route('/pptx-search', methods=['GET','POST'])
def search_function():
   global search_query,search_option
   search_query = request.form.get('searchquery')
   search_option = request.form.get('searchoption')
   pptx_generator_ui.searchvariables(search_query,search_option)
   pptx_generator_ui.search()

    # Prepare the result data
   result = {
        'query': search_query,
        'option': search_option,
        'message': f"Search for '{search_query}' with option '{search_option}' processed successfully.",
        'key': pptx_generator_ui.searchresult_key,
        'author': pptx_generator_ui.searchresult_author,
        'title': pptx_generator_ui.searchresult_title,
        'error': pptx_generator_ui.searchresult_error,
        'invalid': pptx_generator_ui.searchresult_invalid,
    }

   if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return jsonify(result)  # Return JSON for AJAX requests
   else:
      print(result)
        # Render the HTML template for normal form submissions
      return render_template(
            'search.html',
            key=pptx_generator_ui.searchresult_key,
            author=pptx_generator_ui.searchresult_author,
            title=pptx_generator_ui.searchresult_title,
            error=pptx_generator_ui.searchresult_error,
            invalid=pptx_generator_ui.searchresult_invalid
        )

   return render_template('search.html',key = pptx_generator_ui.searchresult_key,author = pptx_generator_ui.searchresult_author,title = pptx_generator_ui.searchresult_title,error = pptx_generator_ui.searchresult_error, invalid = pptx_generator_ui.searchresult_invalid)
@app.route('/generator-function', methods=['GET','POST'])
def generator_function():
   global song_1,song_2,song_3
   if request.method == 'POST':
      song_1 = request.form.get('song1')
      song_2 = request.form.get('song2')
      song_3 = request.form.get('song3')
      presentation_title = request.form.get('pptxtitle')
      if presentation_title == "":
         presentation_title = 'presentation'
      pptx_generator_ui.songvariables(song_1,song_2,song_3,presentation_title)
      pptx_generator_ui.main()

      @after_this_request
      def remove_file(response):
         os.remove(pptx_generator_ui.output_presentation)
         return response
       
      return send_file(pptx_generator_ui.output_presentation,as_attachment=True)
   else:
      return render_template('index.html')
      
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 6000))  # Default to 5000 if PORT not set
    app.run(host='0.0.0.0', port=port, debug=True)
