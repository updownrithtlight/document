from flask import jsonify


class ResponseTemplate:
    @staticmethod
    def success(data=None, message=''):
        return jsonify({'success': True, 'data': data or {}, 'message': message})

    @staticmethod
    def error(message=''):
        return jsonify({'success': False, 'data': {}, 'message': message})


class ResponsePageTemplate:
    @staticmethod
    def success(data=None, message='', total_pages=0, current_page=0):
        return jsonify({'success': True, 'data': data or {}, 'totalPages': total_pages, 'currentPage': current_page,
                        'message': message})

    @staticmethod
    def error(message=''):
        return jsonify({'success': False, 'data': {}, 'message': message})
