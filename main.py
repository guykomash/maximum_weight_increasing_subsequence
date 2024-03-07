from openpyxl import Workbook
import random
import pandas
import time

#  data manipulation

def create_data(n):
    """creating data.xlsx file with n points."""
    wb = Workbook()
    ws = wb.active
    ws.title = "example_data"

    for row in range(1,n + 1):
        random_number_1 = random.randint(1, 1000)
        random_number_2 = random.randint(1, 1000)

        ws[f"A{row}"] = random_number_1
        ws[f"B{row}"] = random_number_2

    try:
        wb.save("data.xlsx")
    except:
        print("Error occured in wb.save()")

def get_points_from_data():
    """returing points from data.xlsx. returning array of sorted points, 
    sorted by x value. if x value is equal, sort by y value"""
    print("getting points from data...")
    try:
      DataFrame = pandas.read_excel('data.xlsx',header=None)
    except:
       print("Error reading xlsx file")
    
    num_of_points = DataFrame.shape[0]
    points = [(DataFrame[0][i], DataFrame[1][i]) for i in range(num_of_points)]
    sorted_points = sorted(points, key=lambda tup: (tup[0], tup[1]))
    return sorted_points

def write_output(weighted_points,squared_weighted_points):
    """creating output.xlsx file with n points."""

    for i,point in enumerate(squared_weighted_points):
      x,y,w = weighted_points[i]
      sq = point[2]
      weighted_points[i] = (x,y,w,sq)

    wb = Workbook()
    ws = wb.active
    ws.title = "output_data"

    for i,point in enumerate(weighted_points):
      x,y,w,sq = point 
      ws[f"A{i+1}"] = x
      ws[f"B{i+1}"] = y
      ws[f"C{i+1}"] = w
      ws[f"D{i+1}"] = sq
    try:
        wb.save("output.xlsx")
    except:
        print("Error occured in wb.save()")

def normalize_points(points):
    """normalize points to"""
    max_x = max(point[0] for point in points)
    min_x = min(point[0] for point in points)
    max_y = max(point[1] for point in points)
    min_y = min(point[1] for point in points)
    normalized_points = []
    for point in points:
        x = (point[0] - min_x) / (max_x - min_x)
        y = (point[1] - min_y) / (max_y - min_y)
        normalized_points.append((x, y))
    return normalized_points

# algorithm

def get_set_of_points_on_heavies_chains(points,weights_map):
    """returns a map: KEY = point (x,y) , VALUE = ( _w , set) : _w is the weight of heaviest chain starting at point. set is a set of all points on _w weight chains starting from point."""
    
    print("mapping points to points on heaviest chains...")

    # KEY = point , VALUE = (_w, set()) . this is the returned map.
    point_set_map = {} 

    """
    Iterate from end to start.
    for each point (point_i) calculate the following:
    1. iset = set of all points that are on one of the heaviest weight chains starting from point_i.
    2. _w = weight of the heaviest chains.
    
    """   
    for i in range(len(points) - 1 , -1 ,-1):
        # for printing. 
        perc = 100 - (i / len(points)) * 100

        point_i = (points[i][0],points[i][1])
        iset = set()
        _w = 0 
        

        # find the max _w belongs to any comparable bigger point (point_j)
        for j in range(i+1,len(points)):
            point_j = (points[j][0],points[j][1])
            if point_j[0] > point_i[0] and point_j[1] > point_i[1]:
              # point_j is a comparable bigger point!
              if (value := point_set_map.get(points[j])) is not None:
                # point_j has been found.
                # value = (point_j_w ,jset)
                jweight = value[0]
                _w = max(_w,jweight)
        
        # now _w is the weight of the heaviest chain aviable from point_i.
        # iterate over point_j again, adding to iset all the jset points only from _w weighted point_js.
        for j in range(i+1, len(points)):
            point_j = (points[j][0],points[j][1])
            # avoid sub-chains by passing points that already in iset.
            if point_j in iset: continue
            if point_j[0] > point_i[0] and point_j[1] > point_i[1]:
              if (value := point_set_map.get(points[j])) is not None:
                
                (jweight, jset) = value
                if _w > jweight : continue

                # point_j meets the conditions...
                # insert to iset all the jset point.
                iset.update(jset)
                iset.add((point_i))
    

        if not iset:
          # this point has no comparable bigger points.
          # update the map to hold an iset of itself with its weight.
          iset = set()
          iset.add(point_i)
          point_set_map[point_i] = (weights_map.get(point_i,1),iset)
        else:
          # update the map: update with new iset and iweight = _w + weight of point_i
          point_set_map[point_i] = (_w + weights_map.get(point_i,1) , iset)
        

        # for printing percentage...
        if perc % 10 == 0:
          if perc != 100: 
            print(f"{perc}%",end=" ",flush=True)
          else: print(f"{perc}%")

    # reutrn the map
    return point_set_map 

def update_points_weights(sorted_points,point_set_map,weights_map,start_w,W):
  """given the points and the point to sets map, this function increase weights of points that starts max weight chain (weighted less than W) by 1.
      call this function after get_set_of_points_on_heavies_chains() to increase the overall weights.
      can be called in iterations to maximize weights."""
  
  print("updating points weights...")

  ending_weight = start_w

  # A flag to identify a increased happend. True will cause another iteration.
  increased = False
  visited = set()
  for point in sorted_points:
      # if point is on a maxed weight chain, continue.
      if point in visited: continue

      _w , point_set = point_set_map[point]
      
      # add point_set to the visited set.
      visited.update(point_set)
      
      # check if point weight can be increased.
      can_increased_by = W - _w
      if can_increased_by > 0:
        # update the weight of the point.
        weights_map[point] = weights_map.get(point,1) + 1
        # update the weight of the heavies chain starting from point, for next iterations.
        point_set_map[point] = (_w + 1, point_set)

        # for printing.
        ending_weight += 1
        
        # identify a increased happend. this will cause another iteration.
        increased = True

  return ending_weight, increased

def get_set_of_points_on_heavies_chains_squared(points,weights_map):
    """returns a map: KEY = point (x,y) , VALUE = ( _w , set) : _w is the weight of heaviest chain starting at point. set is a set of all points on _w weight chains starting from point."""
    
    print("mapping points to points on heaviest chains...")

    # KEY = point , VALUE = (_w, set()) . this is the returned map.
    point_set_map = {} 

    """
    Iterate from end to start.
    for each point (point_i) calculate the following:
    1. iset = set of all points that are on one of the heaviest weight chains starting from point_i.
    2. _w = weight of the heaviest chains.
    
    """   
    for i in range(len(points) - 1 , -1 ,-1):
        # for printing. 
        perc = 100 - (i / len(points)) * 100

        point_i = (points[i][0],points[i][1])
        iset = set()
        _w = 0 
        

        # find the max _w belongs to any comparable bigger point (point_j)
        for j in range(i+1,len(points)):
            point_j = (points[j][0],points[j][1])
            if point_j[0] > point_i[0] and point_j[1] > point_i[1]:
              # point_j is a comparable bigger point!
              if (value := point_set_map.get(points[j])) is not None:
                # point_j has been found.
                # value = (point_j_w ,jset)
                jweight = value[0]
                _w = max(_w,jweight)
        
        # now _w is the weight of the heaviest chain aviable from point_i.
        # iterate over point_j again, adding to iset all the jset points only from _w weighted point_js.
        for j in range(i+1, len(points)):
            point_j = (points[j][0],points[j][1])
            # avoid sub-chains by passing points that already in iset.
            if point_j in iset: continue
            if point_j[0] > point_i[0] and point_j[1] > point_i[1]:
              if (value := point_set_map.get(points[j])) is not None:
                
                (jweight, jset) = value
                curr_weight = weights_map.get(point_i,1)


                # skip all point_j that are not _w weighted or will not get impacted if point_i weight will be increased by one.
                if jweight +  (curr_weight + 1) ** 2 < _w: continue

                # insert to iset all the jset point.
                iset.update(jset)
                iset.add((point_i))
    

        if not iset:
          # this point has no comparable bigger points.
          # update the map to hold an iset of itself with its weight.
          iset = set()
          iset.add(point_i)
          point_set_map[point_i] = (weights_map.get(point_i,1) ** 2,iset)
        else:
          # update the map: update with new iset and iweight = _w + weight of point_i
          point_set_map[point_i] = (_w + weights_map.get(point_i,1) ** 2 , iset)
        

        # for printing percentage...
        if perc % 10 == 0:
          if perc != 100: 
            print(f"{perc}%",end=" ",flush=True)
          else: print(f"{perc}%")

    # reutrn the map
    return point_set_map 

def update_points_weights_squared(sorted_points,point_set_map,weights_map,start_w,W):
  """given the points and the point to sets map, this function increase weights of points that starts max weight chain (weighted less than W) by 1.
      call this function after get_set_of_points_on_heavies_chains() to increase the overall weights.
      can be called in iterations to maximize weights."""
  
  print("updating points weights...")

  ending_weight = start_w

  # A flag to identify a increased happend. True will cause another iteration.
  increased = False
  visited = set()
  for point in sorted_points:
      # if point is on a maxed weight chain, continue.
      if point in visited: continue

      _w , point_set = point_set_map[point]
      
      curr_weight = weights_map.get(point,1)
      new_w = _w - curr_weight + (curr_weight + 1) ** 2

      # add point_set to the visited set.
      visited.update(point_set)
      
      # check if point weight can be increased.
      can_increased_by = W - new_w
      if can_increased_by > 0:
        # update the weight of the point.

        # need to increase point weight by 1.
        # increase _w by one each iteration.
        weights_map[point] = weights_map.get(point,1) + 1
        # update the weight of the heavies chain starting from point, for next iterations.
        point_set_map[point] = (new_w, point_set)

        # for printing.
        ending_weight += 1
        
        # identify a increased happend. this will cause another iteration.
        increased = True

  return ending_weight, increased

def maximum_weight_iteration(points,point_set_map,weights_map,start_w,W,squared):
    if squared:
      point_set_map = get_set_of_points_on_heavies_chains_squared(points,weights_map)
      end_w,is_increased = update_points_weights_squared(points,point_set_map,weights_map,start_w,W)
      return end_w,is_increased
    else:
      point_set_map = get_set_of_points_on_heavies_chains(points,weights_map)
      end_w,is_increased = update_points_weights(points,point_set_map,weights_map,start_w,W)
      return end_w,is_increased

def get_W(points,weights_map):
  n = len(points)
  dp = [0] * n

  weighted_points = []
  for point in points:
    w = weights_map.get(point,1)
    weighted_points.append((point[0],point[1],w))
  for i in range(n - 1, -1, -1):
      xi, yi, wi = weighted_points[i]
      dp[i] = wi

      for j in range(i + 1, n):
          xj, yj, wj = weighted_points[j]

          if xi < xj and yi < yj:
              if dp[i] < dp[j] + wi:
                  dp[i] = dp[j] + wi
  return max(dp),weighted_points


  squared_points = []
  for point in points:
    x = (point[0]*point[0])
    y = (point[1]*point[1])
    squared_points.append((x, y))
  return squared_points

def run(points,squared=False):
    # points = normalize_points(points)

  # ______  saved results for printing.
  all_weights = []

  # ------- initalizition.
  run_iteration = True
  weights_map = {}
  point_set_map = {}
  start_w = len(points)
  i = 0

  # start time
  start_time = time.time()
  
  # ------- run iterations.
  first_W, _ = get_W(points,weights_map)
  while (run_iteration):
    i += 1
    print(f"running iteration [{i}]")
    end_w,run_iteration = maximum_weight_iteration(points,point_set_map,weights_map,start_w,first_W,squared=squared)
  
    # for printing.
    all_weights.append((start_w,end_w))
    start_w = end_w

    # check if excceding max run time.
    time_passed = time.time() - start_time
    if time_passed / 60 > MAX_RUNTIME_MINUTES: run_iteration = False
    print(time_passed)
  
  # _____________________________________________________________________
  
  last_W, weighted_points = get_W(points,weights_map)
  total_time = time.time() - start_time
  print("Iterations Completed!")
  print(f"running time : {total_time}")
  
  # ___________ print results
  print(f"In {i} iterations: total weight increase: {all_weights[0][0]} -> {all_weights[-1][1]}")
  print(f"Max chain weight (W) sanity check : first iteration = {first_W}, after {i} iterations = {last_W} ")

  return weighted_points

if __name__ == "__main__":
  NUM_POINTS = 40000
  MAX_RUNTIME_MINUTES  = 1  # in minutes.
  
  
  create_data(NUM_POINTS)
  points = get_points_from_data()
  points = normalize_points(points)
  
  weighted_points = run(points)
  squared_weighted_points = run(points,squared=True) 

  write_output(weighted_points,squared_weighted_points)

